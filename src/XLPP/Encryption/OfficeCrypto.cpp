#include "OfficeCrypto.h"
#include "../Packaging/AtomicFile.h"

#include "../Vba/VbaProjectBinary.h"

#if defined(_WIN32)
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <bcrypt.h>
#pragma comment(lib, "bcrypt.lib")
#else
#include <openssl/evp.h>
#include <openssl/hmac.h>
#include <openssl/rand.h>
#endif

#include <algorithm>
#include <array>
#include <cctype>
#include <cstdint>
#include <cstring>
#include <fstream>
#include <limits>
#include <map>
#include <memory>
#include <stdexcept>
#include <string_view>

namespace xlpp::internal {
namespace {

constexpr std::size_t kSaltSize = 16;
constexpr std::size_t kBlockSize = 16;
constexpr std::size_t kSegmentSize = 4096;

constexpr std::array<unsigned char, 8> kVerifierHashInputBlockKey{
    0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79};
constexpr std::array<unsigned char, 8> kVerifierHashValueBlockKey{
    0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e};
constexpr std::array<unsigned char, 8> kEncryptedKeyValueBlockKey{
    0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6};
constexpr std::array<unsigned char, 8> kIntegrityKeyBlockKey{
    0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6};
constexpr std::array<unsigned char, 8> kIntegrityValueBlockKey{
    0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33};

using Bytes = std::vector<unsigned char>;

std::size_t roundUp(std::size_t value, std::size_t block) {
    if (!block) return value;
    const auto remainder = value % block;
    return remainder ? value + (block - remainder) : value;
}

void appendU16(Bytes& out, std::uint16_t value) {
    out.push_back(static_cast<unsigned char>(value & 0xffu));
    out.push_back(static_cast<unsigned char>((value >> 8u) & 0xffu));
}

void appendU32(Bytes& out, std::uint32_t value) {
    for (unsigned shift = 0; shift < 32; shift += 8)
        out.push_back(static_cast<unsigned char>((value >> shift) & 0xffu));
}

void appendU64(Bytes& out, std::uint64_t value) {
    for (unsigned shift = 0; shift < 64; shift += 8)
        out.push_back(static_cast<unsigned char>((value >> shift) & 0xffu));
}

std::uint16_t getU16(const unsigned char* p) {
    return static_cast<std::uint16_t>(p[0]) | (static_cast<std::uint16_t>(p[1]) << 8u);
}

std::uint32_t getU32(const unsigned char* p) {
    return static_cast<std::uint32_t>(p[0]) |
           (static_cast<std::uint32_t>(p[1]) << 8u) |
           (static_cast<std::uint32_t>(p[2]) << 16u) |
           (static_cast<std::uint32_t>(p[3]) << 24u);
}

std::uint64_t getU64(const unsigned char* p) {
    std::uint64_t value = 0;
    for (unsigned shift = 0; shift < 64; shift += 8)
        value |= static_cast<std::uint64_t>(p[shift / 8]) << shift;
    return value;
}

#if defined(_WIN32)

void bcryptCheck(NTSTATUS status, const char* operation) {
    if (status < 0) throw std::runtime_error(std::string("Windows CNG failed during ") + operation);
}

struct BCryptAlgorithmHandle {
    BCRYPT_ALG_HANDLE value{nullptr};
    ~BCryptAlgorithmHandle() { if (value) BCryptCloseAlgorithmProvider(value, 0); }
};

struct BCryptHashHandle {
    BCRYPT_HASH_HANDLE value{nullptr};
    ~BCryptHashHandle() { if (value) BCryptDestroyHash(value); }
};

struct BCryptKeyHandle {
    BCRYPT_KEY_HANDLE value{nullptr};
    ~BCryptKeyHandle() { if (value) BCryptDestroyKey(value); }
};

ULONG bcryptPropertyUlong(BCRYPT_HANDLE handle, LPCWSTR name) {
    ULONG value = 0, written = 0;
    bcryptCheck(BCryptGetProperty(handle, name, reinterpret_cast<PUCHAR>(&value), sizeof(value), &written, 0),
                "property query");
    if (written != sizeof(value)) throw std::runtime_error("Windows CNG returned an invalid ULONG property size");
    return value;
}

Bytes bcryptHash(LPCWSTR algorithm, std::initializer_list<std::pair<const unsigned char*, std::size_t>> parts,
                 const Bytes* hmacKey = nullptr) {
    BCryptAlgorithmHandle provider;
    bcryptCheck(BCryptOpenAlgorithmProvider(&provider.value, algorithm, nullptr,
                hmacKey ? BCRYPT_ALG_HANDLE_HMAC_FLAG : 0), "hash provider initialization");
    const auto objectSize = bcryptPropertyUlong(provider.value, BCRYPT_OBJECT_LENGTH);
    const auto hashSizeBytes = bcryptPropertyUlong(provider.value, BCRYPT_HASH_LENGTH);
    std::vector<unsigned char> object(objectSize);
    BCryptHashHandle hash;
    PUCHAR secret = hmacKey && !hmacKey->empty() ? const_cast<PUCHAR>(hmacKey->data()) : nullptr;
    ULONG secretSize = hmacKey ? static_cast<ULONG>(hmacKey->size()) : 0;
    bcryptCheck(BCryptCreateHash(provider.value, &hash.value, object.data(), static_cast<ULONG>(object.size()),
                                 secret, secretSize, 0), "hash creation");
    for (const auto& [data, size] : parts) {
        std::size_t offset = 0;
        while (offset < size) {
            const auto chunk = static_cast<ULONG>(std::min<std::size_t>(size - offset, 1u << 30));
            bcryptCheck(BCryptHashData(hash.value, const_cast<PUCHAR>(data + offset), chunk, 0), "hash update");
            offset += chunk;
        }
    }
    Bytes result(hashSizeBytes);
    bcryptCheck(BCryptFinishHash(hash.value, result.data(), static_cast<ULONG>(result.size()), 0), "hash finalization");
    SecureZeroMemory(object.data(), object.size());
    return result;
}

LPCWSTR bcryptHashAlgorithm(PackageEncryptionHash algorithm) {
    switch (algorithm) {
        case PackageEncryptionHash::Sha1: return BCRYPT_SHA1_ALGORITHM;
        case PackageEncryptionHash::Sha256: return BCRYPT_SHA256_ALGORITHM;
        case PackageEncryptionHash::Sha384: return BCRYPT_SHA384_ALGORITHM;
        case PackageEncryptionHash::Sha512: return BCRYPT_SHA512_ALGORITHM;
    }
    throw std::runtime_error("Unsupported Office encryption hash algorithm");
}

Bytes hashDigest(PackageEncryptionHash algorithm, const Bytes& bytes) {
    return bcryptHash(bcryptHashAlgorithm(algorithm), {{bytes.data(), bytes.size()}});
}

Bytes hashParts(PackageEncryptionHash algorithm,
                std::initializer_list<std::pair<const unsigned char*, std::size_t>> parts) {
    return bcryptHash(bcryptHashAlgorithm(algorithm), parts);
}

Bytes hmacDigest(PackageEncryptionHash algorithm, const Bytes& key, const Bytes& message) {
    return bcryptHash(bcryptHashAlgorithm(algorithm), {{message.data(), message.size()}}, &key);
}

Bytes randomBytes(std::size_t count) {
    Bytes result(count);
    std::size_t offset = 0;
    while (offset < count) {
        const auto chunk = static_cast<ULONG>(std::min<std::size_t>(count - offset, 1u << 30));
        bcryptCheck(BCryptGenRandom(nullptr, result.data() + offset, chunk, BCRYPT_USE_SYSTEM_PREFERRED_RNG),
                    "random-number generation");
        offset += chunk;
    }
    return result;
}

Bytes bcryptAes(const Bytes& input, const Bytes& key, const Bytes* iv, bool encrypt) {
    if ((key.size() != 16 && key.size() != 24 && key.size() != 32) || input.size() % kBlockSize != 0)
        throw std::invalid_argument("Invalid AES key/input geometry");
    if (iv && iv->size() != kBlockSize) throw std::invalid_argument("Invalid AES-CBC IV geometry");
    BCryptAlgorithmHandle provider;
    bcryptCheck(BCryptOpenAlgorithmProvider(&provider.value, BCRYPT_AES_ALGORITHM, nullptr, 0), "AES provider initialization");
    const wchar_t* mode = iv ? BCRYPT_CHAIN_MODE_CBC : BCRYPT_CHAIN_MODE_ECB;
    bcryptCheck(BCryptSetProperty(provider.value, BCRYPT_CHAINING_MODE,
                reinterpret_cast<PUCHAR>(const_cast<wchar_t*>(mode)),
                static_cast<ULONG>((wcslen(mode) + 1) * sizeof(wchar_t)), 0), "AES chaining-mode configuration");
    const auto objectSize = bcryptPropertyUlong(provider.value, BCRYPT_OBJECT_LENGTH);
    std::vector<unsigned char> object(objectSize);
    BCryptKeyHandle keyHandle;
    bcryptCheck(BCryptGenerateSymmetricKey(provider.value, &keyHandle.value, object.data(), static_cast<ULONG>(object.size()),
                const_cast<PUCHAR>(key.data()), static_cast<ULONG>(key.size()), 0), "AES key import");

    Bytes result(input.size());
    if (iv) {
        Bytes mutableIv(*iv);
        ULONG written = 0;
        NTSTATUS status = encrypt
            ? BCryptEncrypt(keyHandle.value, const_cast<PUCHAR>(input.data()), static_cast<ULONG>(input.size()), nullptr,
                            mutableIv.data(), static_cast<ULONG>(mutableIv.size()), result.data(), static_cast<ULONG>(result.size()), &written, 0)
            : BCryptDecrypt(keyHandle.value, const_cast<PUCHAR>(input.data()), static_cast<ULONG>(input.size()), nullptr,
                            mutableIv.data(), static_cast<ULONG>(mutableIv.size()), result.data(), static_cast<ULONG>(result.size()), &written, 0);
        bcryptCheck(status, encrypt ? "AES-CBC encryption" : "AES-CBC decryption");
        result.resize(written);
        SecureZeroMemory(mutableIv.data(), mutableIv.size());
        SecureZeroMemory(object.data(), object.size());
        return result;
    }

    constexpr std::size_t kChunk = 16u * 1024u * 1024u;
    std::size_t offset = 0;
    while (offset < input.size()) {
        const auto count = std::min<std::size_t>(input.size() - offset, kChunk);
        ULONG written = 0;
        NTSTATUS status = encrypt
            ? BCryptEncrypt(keyHandle.value, const_cast<PUCHAR>(input.data() + offset), static_cast<ULONG>(count), nullptr,
                            nullptr, 0, result.data() + offset, static_cast<ULONG>(count), &written, 0)
            : BCryptDecrypt(keyHandle.value, const_cast<PUCHAR>(input.data() + offset), static_cast<ULONG>(count), nullptr,
                            nullptr, 0, result.data() + offset, static_cast<ULONG>(count), &written, 0);
        bcryptCheck(status, encrypt ? "AES-ECB encryption" : "AES-ECB decryption");
        if (written != count) throw std::runtime_error("Windows CNG returned an unexpected AES-ECB output size");
        offset += count;
    }
    SecureZeroMemory(object.data(), object.size());
    return result;
}

Bytes aesCbc(const Bytes& input, const Bytes& key, const Bytes& iv, bool encrypt) {
    return bcryptAes(input, key, &iv, encrypt);
}

Bytes aesEcb(const Bytes& input, const Bytes& key, bool encrypt) {
    return bcryptAes(input, key, nullptr, encrypt);
}

void secureZero(Bytes& bytes) noexcept {
    if (!bytes.empty()) SecureZeroMemory(bytes.data(), bytes.size());
}

#else

const EVP_MD* evpHash(PackageEncryptionHash algorithm) {
    switch (algorithm) {
        case PackageEncryptionHash::Sha1: return EVP_sha1();
        case PackageEncryptionHash::Sha256: return EVP_sha256();
        case PackageEncryptionHash::Sha384: return EVP_sha384();
        case PackageEncryptionHash::Sha512: return EVP_sha512();
    }
    throw std::runtime_error("Unsupported Office encryption hash algorithm");
}

Bytes hashDigest(PackageEncryptionHash algorithm, const Bytes& bytes) {
    const EVP_MD* md = evpHash(algorithm);
    Bytes result(static_cast<std::size_t>(EVP_MD_get_size(md)));
    unsigned int size = 0;
    if (EVP_Digest(bytes.data(), bytes.size(), result.data(), &size, md, nullptr) != 1 || size != result.size())
        throw std::runtime_error("OpenSSL hash operation failed");
    return result;
}

Bytes hashParts(PackageEncryptionHash algorithm,
                std::initializer_list<std::pair<const unsigned char*, std::size_t>> parts) {
    EVP_MD_CTX* raw = EVP_MD_CTX_new();
    if (!raw) throw std::runtime_error("OpenSSL could not allocate hash context");
    std::unique_ptr<EVP_MD_CTX, decltype(&EVP_MD_CTX_free)> ctx(raw, EVP_MD_CTX_free);
    const EVP_MD* md = evpHash(algorithm);
    if (EVP_DigestInit_ex(ctx.get(), md, nullptr) != 1)
        throw std::runtime_error("OpenSSL hash initialization failed");
    for (const auto& [data, size] : parts)
        if (size && EVP_DigestUpdate(ctx.get(), data, size) != 1)
            throw std::runtime_error("OpenSSL hash update failed");
    Bytes result(static_cast<std::size_t>(EVP_MD_get_size(md)));
    unsigned int size = 0;
    if (EVP_DigestFinal_ex(ctx.get(), result.data(), &size) != 1 || size != result.size())
        throw std::runtime_error("OpenSSL hash finalization failed");
    return result;
}

Bytes hmacDigest(PackageEncryptionHash algorithm, const Bytes& key, const Bytes& message) {
    const EVP_MD* md = evpHash(algorithm);
    Bytes result(static_cast<std::size_t>(EVP_MD_get_size(md)));
    unsigned int size = 0;
    if (!HMAC(md, key.data(), static_cast<int>(key.size()), message.data(), message.size(), result.data(), &size)
        || size != result.size())
        throw std::runtime_error("OpenSSL HMAC operation failed");
    return result;
}

Bytes randomBytes(std::size_t count) {
    Bytes result(count);
    if (count && RAND_bytes(result.data(), static_cast<int>(count)) != 1)
        throw std::runtime_error("OpenSSL RAND_bytes failed while generating Office encryption salt/key material");
    return result;
}

const EVP_CIPHER* evpAesCipher(std::size_t keyBytes, bool cbc) {
    if (cbc) {
        if (keyBytes == 16) return EVP_aes_128_cbc();
        if (keyBytes == 24) return EVP_aes_192_cbc();
        if (keyBytes == 32) return EVP_aes_256_cbc();
    } else {
        if (keyBytes == 16) return EVP_aes_128_ecb();
        if (keyBytes == 24) return EVP_aes_192_ecb();
        if (keyBytes == 32) return EVP_aes_256_ecb();
    }
    throw std::invalid_argument("Invalid AES key length");
}

Bytes evpAes(const Bytes& input, const Bytes& key, const Bytes* iv, bool encrypt) {
    if (input.size() % kBlockSize != 0 || (iv && iv->size() != kBlockSize))
        throw std::invalid_argument("Invalid AES input/IV geometry");
    EVP_CIPHER_CTX* raw = EVP_CIPHER_CTX_new();
    if (!raw) throw std::runtime_error("OpenSSL could not allocate AES context");
    std::unique_ptr<EVP_CIPHER_CTX, decltype(&EVP_CIPHER_CTX_free)> ctx(raw, EVP_CIPHER_CTX_free);
    const auto* cipher = evpAesCipher(key.size(), iv != nullptr);
    if (EVP_CipherInit_ex(ctx.get(), cipher, nullptr, key.data(), iv ? iv->data() : nullptr, encrypt ? 1 : 0) != 1)
        throw std::runtime_error("OpenSSL AES initialization failed");
    EVP_CIPHER_CTX_set_padding(ctx.get(), 0);
    Bytes result(input.size() + kBlockSize);
    int first = 0, final = 0;
    if (EVP_CipherUpdate(ctx.get(), result.data(), &first, input.data(), static_cast<int>(input.size())) != 1 ||
        EVP_CipherFinal_ex(ctx.get(), result.data() + first, &final) != 1)
        throw std::runtime_error("OpenSSL AES operation failed");
    result.resize(static_cast<std::size_t>(first + final));
    return result;
}

Bytes aesCbc(const Bytes& input, const Bytes& key, const Bytes& iv, bool encrypt) {
    return evpAes(input, key, &iv, encrypt);
}

Bytes aesEcb(const Bytes& input, const Bytes& key, bool encrypt) {
    return evpAes(input, key, nullptr, encrypt);
}

void secureZero(Bytes& bytes) noexcept {
    if (!bytes.empty()) OPENSSL_cleanse(bytes.data(), bytes.size());
}

#endif

std::size_t officeHashSize(PackageEncryptionHash algorithm) {
    switch (algorithm) {
        case PackageEncryptionHash::Sha1: return 20;
        case PackageEncryptionHash::Sha256: return 32;
        case PackageEncryptionHash::Sha384: return 48;
        case PackageEncryptionHash::Sha512: return 64;
    }
    throw std::runtime_error("Unsupported Office encryption hash algorithm");
}

const char* officeHashName(PackageEncryptionHash algorithm) {
    switch (algorithm) {
        case PackageEncryptionHash::Sha1: return "SHA-1";
        case PackageEncryptionHash::Sha256: return "SHA256";
        case PackageEncryptionHash::Sha384: return "SHA384";
        case PackageEncryptionHash::Sha512: return "SHA512";
    }
    throw std::runtime_error("Unsupported Office encryption hash algorithm");
}

PackageEncryptionHash parseOfficeHashName(const std::string& name) {
    if (name == "SHA-1" || name == "SHA1") return PackageEncryptionHash::Sha1;
    if (name == "SHA256" || name == "SHA-256") return PackageEncryptionHash::Sha256;
    if (name == "SHA384" || name == "SHA-384") return PackageEncryptionHash::Sha384;
    if (name == "SHA512" || name == "SHA-512") return PackageEncryptionHash::Sha512;
    throw std::runtime_error("Unsupported Agile Encryption hash algorithm: " + name);
}

bool constantTimeEqual(const unsigned char* a, const unsigned char* b, std::size_t count) noexcept {
    unsigned char diff = 0;
    for (std::size_t i = 0; i < count; ++i) diff |= static_cast<unsigned char>(a[i] ^ b[i]);
    return diff == 0;
}

struct WipeOnExit {
    Bytes* bytes{};
    ~WipeOnExit() { if (bytes) secureZero(*bytes); }
};


Bytes utf8ToUtf16Le(std::string_view text) {
    Bytes result;
    result.reserve(text.size() * 2);
    std::size_t i = 0;
    while (i < text.size()) {
        const auto first = static_cast<unsigned char>(text[i]);
        std::uint32_t cp = 0;
        std::size_t count = 0;
        if (first < 0x80) { cp = first; count = 1; }
        else if ((first & 0xe0u) == 0xc0u) { cp = first & 0x1fu; count = 2; }
        else if ((first & 0xf0u) == 0xe0u) { cp = first & 0x0fu; count = 3; }
        else if ((first & 0xf8u) == 0xf0u) { cp = first & 0x07u; count = 4; }
        else throw std::invalid_argument("Password contains invalid UTF-8");
        if (i + count > text.size()) throw std::invalid_argument("Password contains truncated UTF-8");
        for (std::size_t j = 1; j < count; ++j) {
            const auto ch = static_cast<unsigned char>(text[i + j]);
            if ((ch & 0xc0u) != 0x80u) throw std::invalid_argument("Password contains invalid UTF-8 continuation byte");
            cp = (cp << 6u) | (ch & 0x3fu);
        }
        const std::uint32_t minValue = count == 1 ? 0 : count == 2 ? 0x80u : count == 3 ? 0x800u : 0x10000u;
        if (cp < minValue || cp > 0x10ffffu || (cp >= 0xd800u && cp <= 0xdfffu))
            throw std::invalid_argument("Password contains invalid UTF-8 code point");
        auto append16 = [&](std::uint16_t value) {
            result.push_back(static_cast<unsigned char>(value & 0xffu));
            result.push_back(static_cast<unsigned char>((value >> 8u) & 0xffu));
        };
        if (cp <= 0xffffu) append16(static_cast<std::uint16_t>(cp));
        else {
            cp -= 0x10000u;
            append16(static_cast<std::uint16_t>(0xd800u + (cp >> 10u)));
            append16(static_cast<std::uint16_t>(0xdc00u + (cp & 0x3ffu)));
        }
        i += count;
    }
    return result;
}

Bytes derivePasswordKey(const std::string& password, const Bytes& salt, std::uint32_t spinCount,
                        PackageEncryptionHash hashAlgorithm,
                        const unsigned char* blockKey, std::size_t blockKeySize, std::size_t keyBytes) {
    auto passwordBytes = utf8ToUtf16Le(password);
    WipeOnExit wipePassword{&passwordBytes};
    auto hash = hashParts(hashAlgorithm, {{salt.data(), salt.size()}, {passwordBytes.data(), passwordBytes.size()}});
    WipeOnExit wipeHash{&hash};
    for (std::uint32_t i = 0; i < spinCount; ++i) {
        const std::array<unsigned char, 4> iterator{
            static_cast<unsigned char>(i & 0xffu),
            static_cast<unsigned char>((i >> 8u) & 0xffu),
            static_cast<unsigned char>((i >> 16u) & 0xffu),
            static_cast<unsigned char>((i >> 24u) & 0xffu)};
        auto nextHash = hashParts(hashAlgorithm, {{iterator.data(), iterator.size()}, {hash.data(), hash.size()}});
        secureZero(hash);
        hash.swap(nextHash);
        secureZero(nextHash);
    }
    auto finalHash = hashParts(hashAlgorithm, {{hash.data(), hash.size()}, {blockKey, blockKeySize}});
    secureZero(hash);
    hash.swap(finalHash);
    secureZero(finalHash);
    if (hash.size() < keyBytes) hash.resize(keyBytes, 0x36);
    else hash.resize(keyBytes);
    Bytes result = hash;
    return result;
}

Bytes deriveIv(const Bytes& keySalt, PackageEncryptionHash hashAlgorithm,
               const unsigned char* blockKey, std::size_t blockKeySize, std::size_t blockSize) {
    Bytes iv;
    if (blockKey && blockKeySize) iv = hashParts(hashAlgorithm, {{keySalt.data(), keySalt.size()}, {blockKey, blockKeySize}});
    else iv = keySalt;
    if (iv.size() < blockSize) iv.resize(blockSize, 0x36);
    else iv.resize(blockSize);
    return iv;
}

Bytes deriveStandardKey(const std::string& password, const Bytes& salt, std::size_t keyBytes) {
    auto passwordBytes = utf8ToUtf16Le(password);
    WipeOnExit wipePassword{&passwordBytes};
    Bytes seed;
    seed.reserve(salt.size() + passwordBytes.size());
    seed.insert(seed.end(), salt.begin(), salt.end());
    seed.insert(seed.end(), passwordBytes.begin(), passwordBytes.end());
    WipeOnExit wipeSeed{&seed};
    auto hash = hashDigest(PackageEncryptionHash::Sha1, seed);
    WipeOnExit wipeHash{&hash};
    for (std::uint32_t i = 0; i < 50000; ++i) {
        Bytes iteration(4 + hash.size());
        iteration[0] = static_cast<unsigned char>(i & 0xffu);
        iteration[1] = static_cast<unsigned char>((i >> 8u) & 0xffu);
        iteration[2] = static_cast<unsigned char>((i >> 16u) & 0xffu);
        iteration[3] = static_cast<unsigned char>((i >> 24u) & 0xffu);
        std::copy(hash.begin(), hash.end(), iteration.begin() + 4);
        auto nextHash = hashDigest(PackageEncryptionHash::Sha1, iteration);
        secureZero(iteration);
        secureZero(hash);
        hash.swap(nextHash);
        secureZero(nextHash);
    }
    Bytes finalInput(hash);
    WipeOnExit wipeFinalInput{&finalInput};
    finalInput.insert(finalInput.end(), {0, 0, 0, 0});
    auto finalHash = hashDigest(PackageEncryptionHash::Sha1, finalInput);
    WipeOnExit wipeFinalHash{&finalHash};
    Bytes pad1(64, 0x36), pad2(64, 0x5c);
    WipeOnExit wipePad1{&pad1}, wipePad2{&pad2};
    for (std::size_t i = 0; i < finalHash.size(); ++i) { pad1[i] ^= finalHash[i]; pad2[i] ^= finalHash[i]; }
    auto x1 = hashDigest(PackageEncryptionHash::Sha1, pad1);
    auto x2 = hashDigest(PackageEncryptionHash::Sha1, pad2);
    WipeOnExit wipeX1{&x1}, wipeX2{&x2};
    x1.insert(x1.end(), x2.begin(), x2.end());
    if (keyBytes > x1.size()) throw std::runtime_error("Standard Encryption AES key length exceeds derivation output");
    x1.resize(keyBytes);
    Bytes result = x1;
    return result;
}

Bytes zeroPad(Bytes bytes, std::size_t blockSize) {
    bytes.resize(roundUp(bytes.size(), blockSize), 0);
    return bytes;
}

std::string base64Encode(const Bytes& bytes) {
    static constexpr char alphabet[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    std::string out;
    out.reserve(4 * ((bytes.size() + 2) / 3));
    for (std::size_t i = 0; i < bytes.size(); i += 3) {
        const std::uint32_t a = bytes[i];
        const std::uint32_t b = i + 1 < bytes.size() ? bytes[i + 1] : 0;
        const std::uint32_t c = i + 2 < bytes.size() ? bytes[i + 2] : 0;
        const std::uint32_t triple = (a << 16u) | (b << 8u) | c;
        out.push_back(alphabet[(triple >> 18u) & 0x3fu]);
        out.push_back(alphabet[(triple >> 12u) & 0x3fu]);
        out.push_back(i + 1 < bytes.size() ? alphabet[(triple >> 6u) & 0x3fu] : '=');
        out.push_back(i + 2 < bytes.size() ? alphabet[triple & 0x3fu] : '=');
    }
    return out;
}

Bytes base64Decode(std::string text) {
    text.erase(std::remove_if(text.begin(), text.end(), [](unsigned char ch) { return std::isspace(ch) != 0; }), text.end());
    if (text.empty()) return {};
    if (text.size() % 4 != 0) throw std::runtime_error("Malformed base64 in EncryptionInfo");
    auto decode = [](char ch) -> int {
        if (ch >= 'A' && ch <= 'Z') return ch - 'A';
        if (ch >= 'a' && ch <= 'z') return ch - 'a' + 26;
        if (ch >= '0' && ch <= '9') return ch - '0' + 52;
        if (ch == '+') return 62;
        if (ch == '/') return 63;
        if (ch == '=') return -2;
        return -1;
    };
    Bytes out;
    out.reserve((text.size() / 4) * 3);
    for (std::size_t i = 0; i < text.size(); i += 4) {
        const int a = decode(text[i]), b = decode(text[i + 1]), c = decode(text[i + 2]), d = decode(text[i + 3]);
        if (a < 0 || b < 0 || c == -1 || d == -1 || (c == -2 && d != -2) || (i + 4 != text.size() && (c == -2 || d == -2)))
            throw std::runtime_error("Malformed base64 in EncryptionInfo");
        const std::uint32_t triple = (static_cast<std::uint32_t>(a) << 18u) |
                                     (static_cast<std::uint32_t>(b) << 12u) |
                                     (static_cast<std::uint32_t>(c < 0 ? 0 : c) << 6u) |
                                     static_cast<std::uint32_t>(d < 0 ? 0 : d);
        out.push_back(static_cast<unsigned char>((triple >> 16u) & 0xffu));
        if (c != -2) out.push_back(static_cast<unsigned char>((triple >> 8u) & 0xffu));
        if (d != -2) out.push_back(static_cast<unsigned char>(triple & 0xffu));
    }
    return out;
}

std::string elementByLocalName(const std::string& xml, const std::string& localName, std::size_t from = 0) {
    std::size_t pos = from;
    while ((pos = xml.find('<', pos)) != std::string::npos) {
        if (pos + 1 < xml.size() && (xml[pos + 1] == '/' || xml[pos + 1] == '?' || xml[pos + 1] == '!')) { ++pos; continue; }
        auto nameEnd = pos + 1;
        while (nameEnd < xml.size() && !std::isspace(static_cast<unsigned char>(xml[nameEnd])) && xml[nameEnd] != '>' && xml[nameEnd] != '/') ++nameEnd;
        const auto fullName = xml.substr(pos + 1, nameEnd - pos - 1);
        const auto colon = fullName.rfind(':');
        const auto local = colon == std::string::npos ? fullName : fullName.substr(colon + 1);
        if (local == localName) {
            const auto end = xml.find('>', nameEnd);
            if (end == std::string::npos) break;
            return xml.substr(pos, end - pos + 1);
        }
        pos = nameEnd;
    }
    return {};
}

std::string attr(const std::string& element, const std::string& name) {
    std::size_t pos = 0;
    while ((pos = element.find(name, pos)) != std::string::npos) {
        const bool leftOk = pos == 0 || std::isspace(static_cast<unsigned char>(element[pos - 1])) || element[pos - 1] == '<';
        const auto after = pos + name.size();
        if (!leftOk || after >= element.size() || element[after] != '=') { pos = after; continue; }
        if (after + 1 >= element.size()) break;
        const char quote = element[after + 1];
        if (quote != '\'' && quote != '"') { pos = after + 1; continue; }
        const auto end = element.find(quote, after + 2);
        if (end == std::string::npos) break;
        return element.substr(after + 2, end - after - 2);
    }
    return {};
}

std::uint32_t uintAttr(const std::string& element, const std::string& name) {
    const auto text = attr(element, name);
    if (text.empty()) throw std::runtime_error("EncryptionInfo is missing attribute: " + name);
    const auto value = std::stoull(text);
    if (value > std::numeric_limits<std::uint32_t>::max()) throw std::runtime_error("EncryptionInfo integer is out of range: " + name);
    return static_cast<std::uint32_t>(value);
}

struct AgileKeyEncryptorXml {
    std::string uri;
    std::string encryptedKeyElement;
};

std::vector<AgileKeyEncryptorXml> agileKeyEncryptors(const std::string& xml) {
    std::vector<AgileKeyEncryptorXml> result;
    std::size_t pos = 0;
    while ((pos = xml.find('<', pos)) != std::string::npos) {
        if (pos + 1 >= xml.size() || xml[pos + 1] == '/' || xml[pos + 1] == '?' || xml[pos + 1] == '!') { ++pos; continue; }
        auto nameEnd = pos + 1;
        while (nameEnd < xml.size() && !std::isspace(static_cast<unsigned char>(xml[nameEnd]))
               && xml[nameEnd] != '>' && xml[nameEnd] != '/') ++nameEnd;
        if (nameEnd <= pos + 1) { ++pos; continue; }
        const auto fullName = xml.substr(pos + 1, nameEnd - pos - 1);
        const auto colon = fullName.rfind(':');
        const auto local = colon == std::string::npos ? fullName : fullName.substr(colon + 1);
        if (local != "keyEncryptor") { pos = nameEnd; continue; }
        const auto openEnd = xml.find('>', nameEnd);
        if (openEnd == std::string::npos) break;
        const auto opening = xml.substr(pos, openEnd - pos + 1);
        const auto closeToken = "</" + fullName + ">";
        const auto close = xml.find(closeToken, openEnd + 1);
        if (close == std::string::npos) { pos = openEnd + 1; continue; }
        const auto block = xml.substr(pos, close + closeToken.size() - pos);
        result.push_back({attr(opening, "uri"), elementByLocalName(block, "encryptedKey")});
        pos = close + closeToken.size();
    }
    return result;
}

struct StandardInfo {
    std::size_t keyBytes{};
    Bytes salt;
    Bytes encryptedVerifier;
    std::uint32_t verifierHashSize{};
    Bytes encryptedVerifierHash;
};

StandardInfo parseStandardInfo(const Bytes& stream) {
    if (stream.size() < 12) throw std::runtime_error("Standard EncryptionInfo stream is truncated");
    const auto major = getU16(stream.data());
    const auto minor = getU16(stream.data() + 2);
    if (minor != 2 || (major != 2 && major != 3 && major != 4))
        throw std::runtime_error("Unsupported Standard Encryption version");
    const auto flags = getU32(stream.data() + 4);
    const auto headerSize = static_cast<std::size_t>(getU32(stream.data() + 8));
    if ((flags & 0x24u) != 0x24u || headerSize < 32 || 12 + headerSize + 40 > stream.size())
        throw std::runtime_error("Unsupported or malformed Standard Encryption header");
    const auto* header = stream.data() + 12;
    const auto headerFlags = getU32(header);
    const auto algId = getU32(header + 8);
    const auto algIdHash = getU32(header + 12);
    const auto keyBits = getU32(header + 16);
    if ((headerFlags & 0x24u) != 0x24u || algIdHash != 0x8004u)
        throw std::runtime_error("XL++ Standard Encryption reader requires CryptoAPI AES with SHA-1");
    std::size_t keyBytes = 0;
    if (algId == 0x660eu && keyBits == 128) keyBytes = 16;
    else if (algId == 0x660fu && keyBits == 192) keyBytes = 24;
    else if (algId == 0x6610u && keyBits == 256) keyBytes = 32;
    else throw std::runtime_error("Unsupported Standard Encryption AES key size");
    const auto verifierOffset = 12 + headerSize;
    const auto saltSize = getU32(stream.data() + verifierOffset);
    if (saltSize != 16 || verifierOffset + 40 > stream.size())
        throw std::runtime_error("Malformed Standard Encryption verifier");
    StandardInfo info;
    info.keyBytes = keyBytes;
    info.salt.assign(stream.begin() + static_cast<std::ptrdiff_t>(verifierOffset + 4),
                     stream.begin() + static_cast<std::ptrdiff_t>(verifierOffset + 20));
    info.encryptedVerifier.assign(stream.begin() + static_cast<std::ptrdiff_t>(verifierOffset + 20),
                                  stream.begin() + static_cast<std::ptrdiff_t>(verifierOffset + 36));
    info.verifierHashSize = getU32(stream.data() + verifierOffset + 36);
    if (info.verifierHashSize != 20 || verifierOffset + 72 > stream.size())
        throw std::runtime_error("Unsupported Standard Encryption verifier hash size");
    info.encryptedVerifierHash.assign(stream.begin() + static_cast<std::ptrdiff_t>(verifierOffset + 40),
                                      stream.begin() + static_cast<std::ptrdiff_t>(verifierOffset + 72));
    return info;
}

Bytes decryptStandardPackage(const Bytes& encryptionInfo, const Bytes& encryptedPackage, const std::string& password,
                             const OfficeDecryptionLimits& limits) {
    const auto info = parseStandardInfo(encryptionInfo);
    auto key = deriveStandardKey(password, info.salt, info.keyBytes);
    WipeOnExit wipeKey{&key};
    const auto verifier = aesEcb(info.encryptedVerifier, key, false);
    const auto verifierHash = aesEcb(info.encryptedVerifierHash, key, false);
    const Bytes verifierValue(verifier.begin(), verifier.begin() + 16);
    const auto actualHash = hashDigest(PackageEncryptionHash::Sha1, verifierValue);
    if (verifierHash.size() < actualHash.size() || !constantTimeEqual(verifierHash.data(), actualHash.data(), actualHash.size()))
        throw std::invalid_argument("Incorrect password for encrypted Office workbook");
    if (encryptedPackage.size() < 8 || (encryptedPackage.size() - 8) % kBlockSize != 0)
        throw std::runtime_error("Standard EncryptedPackage stream is malformed");
    const auto originalSize = getU64(encryptedPackage.data());
    if (originalSize > static_cast<std::uint64_t>(std::numeric_limits<std::size_t>::max()))
        throw std::runtime_error("EncryptedPackage is too large for this platform");
    const auto cipherBytes = encryptedPackage.size() - 8;
    if (originalSize > cipherBytes)
        throw std::runtime_error("Standard EncryptedPackage size exceeds ciphertext geometry");
    if (limits.maxPlainPackageBytes != 0 && originalSize > limits.maxPlainPackageBytes)
        throw std::runtime_error("Decrypted workbook exceeds configured maxDecryptedPackageBytes limit");
    Bytes cipher(encryptedPackage.begin() + 8, encryptedPackage.end());
    auto plain = aesEcb(cipher, key, false);
    if (originalSize > plain.size()) throw std::runtime_error("Standard EncryptedPackage plaintext is truncated");
    plain.resize(static_cast<std::size_t>(originalSize));
    return plain;
}

struct AgileAlgorithmInfo {
    std::uint32_t keyBits{};
    std::uint32_t blockSize{};
    std::uint32_t hashSize{};
    PackageEncryptionHash hashAlgorithm{PackageEncryptionHash::Sha512};
    Bytes salt;
    std::string cipherAlgorithm;
    std::string cipherChaining;
};

struct AgileInfo {
    std::uint32_t spinCount{};
    AgileAlgorithmInfo keyData;
    AgileAlgorithmInfo passwordKey;
    Bytes encryptedVerifierInput;
    Bytes encryptedVerifierHash;
    Bytes encryptedKey;
    Bytes encryptedHmacKey;
    Bytes encryptedHmacValue;
};

AgileAlgorithmInfo parseAgileAlgorithm(const std::string& element, bool requireSpin = false) {
    AgileAlgorithmInfo result;
    result.cipherAlgorithm = attr(element, "cipherAlgorithm");
    result.cipherChaining = attr(element, "cipherChaining");
    result.blockSize = uintAttr(element, "blockSize");
    result.keyBits = uintAttr(element, "keyBits");
    result.hashSize = uintAttr(element, "hashSize");
    result.hashAlgorithm = parseOfficeHashName(attr(element, "hashAlgorithm"));
    const auto saltSize = uintAttr(element, "saltSize");
    result.salt = base64Decode(attr(element, "saltValue"));
    if (result.cipherAlgorithm != "AES" || result.cipherChaining != "ChainingModeCBC")
        throw std::runtime_error("XL++ Agile Encryption reader requires AES-CBC");
    if (result.blockSize != kBlockSize)
        throw std::runtime_error("XL++ Agile Encryption reader requires a 16-byte AES block size");
    if (result.keyBits != 128 && result.keyBits != 192 && result.keyBits != 256)
        throw std::runtime_error("Unsupported Agile Encryption AES key size");
    if (result.hashSize != officeHashSize(result.hashAlgorithm))
        throw std::runtime_error("Agile Encryption hashSize does not match hashAlgorithm");
    if (saltSize == 0 || saltSize > 65536u || result.salt.size() != saltSize)
        throw std::runtime_error("Malformed Agile Encryption salt");
    (void)requireSpin;
    return result;
}

AgileInfo parseAgileInfo(const Bytes& stream) {
    if (stream.size() < 8) throw std::runtime_error("EncryptionInfo stream is truncated");
    if (getU16(stream.data()) != 4 || getU16(stream.data() + 2) != 4 || getU32(stream.data() + 4) != 0x40)
        throw std::runtime_error("Only ECMA-376 Agile Encryption version 4.4 is supported");
    const std::string xml(stream.begin() + 8, stream.end());
    const auto keyData = elementByLocalName(xml, "keyData");
    const auto integrity = elementByLocalName(xml, "dataIntegrity");
    const auto keyEncryptors = agileKeyEncryptors(xml);
    const auto passwordUri = std::string("http://schemas.microsoft.com/office/2006/keyEncryptor/password");
    const auto passwordCount = static_cast<std::size_t>(std::count_if(
        keyEncryptors.begin(), keyEncryptors.end(), [&](const auto& item) { return item.uri == passwordUri; }));
    if (passwordCount != 1)
        throw std::runtime_error("Agile EncryptionInfo must contain exactly one password key-encryptor");
    const auto passwordIt = std::find_if(keyEncryptors.begin(), keyEncryptors.end(), [&](const auto& item) {
        return item.uri == passwordUri;
    });
    if (keyData.empty() || passwordIt == keyEncryptors.end() || passwordIt->encryptedKeyElement.empty())
        throw std::runtime_error("Agile EncryptionInfo is missing its password key-encryptor");
    const auto& encryptedKey = passwordIt->encryptedKeyElement;
    AgileInfo info;
    info.keyData = parseAgileAlgorithm(keyData);
    info.passwordKey = parseAgileAlgorithm(encryptedKey, true);
    info.spinCount = uintAttr(encryptedKey, "spinCount");
    if (info.spinCount > 10000000u) throw std::runtime_error("Agile Encryption spinCount exceeds specification limit");
    info.encryptedVerifierInput = base64Decode(attr(encryptedKey, "encryptedVerifierHashInput"));
    info.encryptedVerifierHash = base64Decode(attr(encryptedKey, "encryptedVerifierHashValue"));
    info.encryptedKey = base64Decode(attr(encryptedKey, "encryptedKeyValue"));
    if (!integrity.empty()) {
        info.encryptedHmacKey = base64Decode(attr(integrity, "encryptedHmacKey"));
        info.encryptedHmacValue = base64Decode(attr(integrity, "encryptedHmacValue"));
    }
    if (info.encryptedVerifierInput.empty() || info.encryptedVerifierInput.size() % info.passwordKey.blockSize != 0 ||
        info.encryptedVerifierHash.empty() || info.encryptedVerifierHash.size() % info.passwordKey.blockSize != 0 ||
        info.encryptedKey.empty() || info.encryptedKey.size() % info.passwordKey.blockSize != 0)
        throw std::runtime_error("Malformed Agile password verifier/key ciphertext");
    if ((!info.encryptedHmacKey.empty() && info.encryptedHmacKey.size() % info.keyData.blockSize != 0) ||
        (!info.encryptedHmacValue.empty() && info.encryptedHmacValue.size() % info.keyData.blockSize != 0))
        throw std::runtime_error("Malformed Agile DataIntegrity ciphertext");
    return info;
}

Bytes decryptIntermediateKey(const AgileInfo& info, const std::string& password) {
    const auto passwordKeyBytes = static_cast<std::size_t>(info.passwordKey.keyBits / 8u);
    const auto packageKeyBytes = static_cast<std::size_t>(info.keyData.keyBits / 8u);
    auto verifierKey = derivePasswordKey(password, info.passwordKey.salt, info.spinCount, info.passwordKey.hashAlgorithm,
        kVerifierHashInputBlockKey.data(), kVerifierHashInputBlockKey.size(), passwordKeyBytes);
    auto hashKey = derivePasswordKey(password, info.passwordKey.salt, info.spinCount, info.passwordKey.hashAlgorithm,
        kVerifierHashValueBlockKey.data(), kVerifierHashValueBlockKey.size(), passwordKeyBytes);
    WipeOnExit wipeVerifierKey{&verifierKey}, wipeHashKey{&hashKey};
    const auto iv = deriveIv(info.passwordKey.salt, info.passwordKey.hashAlgorithm, nullptr, 0, info.passwordKey.blockSize);
    const auto verifier = aesCbc(info.encryptedVerifierInput, verifierKey, iv, false);
    const auto expectedHashPadded = aesCbc(info.encryptedVerifierHash, hashKey, iv, false);
    const auto verifierSize = std::min<std::size_t>(info.passwordKey.salt.size(), verifier.size());
    const Bytes verifierActual(verifier.begin(), verifier.begin() + static_cast<std::ptrdiff_t>(verifierSize));
    const auto actualHash = hashDigest(info.passwordKey.hashAlgorithm, verifierActual);
    if (expectedHashPadded.size() < actualHash.size() ||
        !constantTimeEqual(expectedHashPadded.data(), actualHash.data(), actualHash.size()))
        throw std::invalid_argument("Incorrect password for encrypted Office workbook");
    auto keyKey = derivePasswordKey(password, info.passwordKey.salt, info.spinCount, info.passwordKey.hashAlgorithm,
        kEncryptedKeyValueBlockKey.data(), kEncryptedKeyValueBlockKey.size(), passwordKeyBytes);
    WipeOnExit wipeKeyKey{&keyKey};
    auto intermediate = aesCbc(info.encryptedKey, keyKey, iv, false);
    if (intermediate.size() < packageKeyBytes) throw std::runtime_error("Encrypted intermediate key is truncated");
    intermediate.resize(packageKeyBytes);
    return intermediate;
}

Bytes encryptPackageSegments(const Bytes& package, const Bytes& key, const Bytes& salt,
                             PackageEncryptionHash hashAlgorithm, std::size_t blockSize) {
    Bytes result;
    appendU64(result, static_cast<std::uint64_t>(package.size()));
    std::size_t offset = 0;
    std::uint32_t segment = 0;
    while (offset < package.size()) {
        const auto count = std::min(kSegmentSize, package.size() - offset);
        Bytes plain(package.begin() + static_cast<std::ptrdiff_t>(offset), package.begin() + static_cast<std::ptrdiff_t>(offset + count));
        plain = zeroPad(std::move(plain), blockSize);
        std::array<unsigned char, 4> blockKey{
            static_cast<unsigned char>(segment & 0xffu),
            static_cast<unsigned char>((segment >> 8u) & 0xffu),
            static_cast<unsigned char>((segment >> 16u) & 0xffu),
            static_cast<unsigned char>((segment >> 24u) & 0xffu)};
        const auto iv = deriveIv(salt, hashAlgorithm, blockKey.data(), blockKey.size(), blockSize);
        const auto encrypted = aesCbc(plain, key, iv, true);
        result.insert(result.end(), encrypted.begin(), encrypted.end());
        offset += count;
        ++segment;
    }
    return result;
}

Bytes decryptPackageSegments(const Bytes& encryptedPackage, const Bytes& key, const Bytes& salt,
                             PackageEncryptionHash hashAlgorithm, std::size_t blockSize,
                             const OfficeDecryptionLimits& limits) {
    if (encryptedPackage.size() < 8) throw std::runtime_error("EncryptedPackage stream is truncated");
    const auto originalSize = getU64(encryptedPackage.data());
    if (originalSize > static_cast<std::uint64_t>(std::numeric_limits<std::size_t>::max()))
        throw std::runtime_error("EncryptedPackage is too large for this platform");
    const auto cipherBytes = encryptedPackage.size() - 8;
    if (originalSize > cipherBytes)
        throw std::runtime_error("Agile EncryptedPackage size exceeds ciphertext geometry");
    if (limits.maxPlainPackageBytes != 0 && originalSize > limits.maxPlainPackageBytes)
        throw std::runtime_error("Decrypted workbook exceeds configured maxDecryptedPackageBytes limit");
    Bytes result;
    result.reserve(static_cast<std::size_t>(originalSize));
    std::size_t encryptedOffset = 8;
    std::uint64_t plainOffset = 0;
    std::uint32_t segment = 0;
    while (plainOffset < originalSize) {
        const auto plainCount = static_cast<std::size_t>(std::min<std::uint64_t>(kSegmentSize, originalSize - plainOffset));
        const auto cipherCount = roundUp(plainCount, blockSize);
        if (encryptedOffset + cipherCount > encryptedPackage.size()) throw std::runtime_error("EncryptedPackage segment is truncated");
        Bytes cipher(encryptedPackage.begin() + static_cast<std::ptrdiff_t>(encryptedOffset),
                     encryptedPackage.begin() + static_cast<std::ptrdiff_t>(encryptedOffset + cipherCount));
        std::array<unsigned char, 4> blockKey{
            static_cast<unsigned char>(segment & 0xffu),
            static_cast<unsigned char>((segment >> 8u) & 0xffu),
            static_cast<unsigned char>((segment >> 16u) & 0xffu),
            static_cast<unsigned char>((segment >> 24u) & 0xffu)};
        const auto iv = deriveIv(salt, hashAlgorithm, blockKey.data(), blockKey.size(), blockSize);
        const auto plain = aesCbc(cipher, key, iv, false);
        result.insert(result.end(), plain.begin(), plain.begin() + static_cast<std::ptrdiff_t>(plainCount));
        encryptedOffset += cipherCount;
        plainOffset += plainCount;
        ++segment;
    }
    return result;
}

std::string agileXml(std::uint32_t spinCount, std::uint32_t keyBits, PackageEncryptionHash hashAlgorithm,
                     const Bytes& keyDataSalt, const Bytes& passwordSalt,
                     const Bytes& encryptedVerifierInput, const Bytes& encryptedVerifierHash,
                     const Bytes& encryptedKey, const Bytes& encryptedHmacKey, const Bytes& encryptedHmacValue) {
    const auto hashSize = officeHashSize(hashAlgorithm);
    const auto hashName = officeHashName(hashAlgorithm);
    std::string xml;
    xml.reserve(1900);
    xml += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    xml += "<encryption xmlns=\"http://schemas.microsoft.com/office/2006/encryption\" xmlns:p=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">";
    xml += "<keyData saltSize=\"" + std::to_string(keyDataSalt.size()) + "\" blockSize=\"16\" keyBits=\"" + std::to_string(keyBits)
        + "\" hashSize=\"" + std::to_string(hashSize) + "\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\""
        + hashName + "\" saltValue=\"" + base64Encode(keyDataSalt) + "\"/>";
    xml += "<dataIntegrity encryptedHmacKey=\"" + base64Encode(encryptedHmacKey) + "\" encryptedHmacValue=\"" + base64Encode(encryptedHmacValue) + "\"/>";
    xml += "<keyEncryptors><keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\"><p:encryptedKey";
    xml += " spinCount=\"" + std::to_string(spinCount) + "\" saltSize=\"" + std::to_string(passwordSalt.size())
        + "\" blockSize=\"16\" keyBits=\"" + std::to_string(keyBits) + "\" hashSize=\"" + std::to_string(hashSize)
        + "\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"" + hashName + "\"";
    xml += " saltValue=\"" + base64Encode(passwordSalt) + "\" encryptedVerifierHashInput=\"" + base64Encode(encryptedVerifierInput)
        + "\" encryptedVerifierHashValue=\"" + base64Encode(encryptedVerifierHash) + "\" encryptedKeyValue=\"" + base64Encode(encryptedKey) + "\"/>";
    xml += "</keyEncryptor></keyEncryptors></encryption>";
    return xml;
}

std::uint32_t standardAlgId(std::uint32_t keyBits) {
    if (keyBits == 128) return 0x660eu;
    if (keyBits == 192) return 0x660fu;
    if (keyBits == 256) return 0x6610u;
    throw std::invalid_argument("Standard Encryption keyBits must be 128, 192, or 256");
}

Bytes utf16LeNullTerminated(std::string_view ascii) {
    Bytes out;
    out.reserve((ascii.size() + 1) * 2);
    for (const unsigned char ch : ascii) { out.push_back(ch); out.push_back(0); }
    out.push_back(0); out.push_back(0);
    return out;
}

Bytes standardEncryptionInfo(const Bytes& salt, const Bytes& encryptedVerifier,
                             const Bytes& encryptedVerifierHash, std::uint32_t keyBits) {
    static constexpr std::string_view kProvider = "Microsoft Enhanced RSA and AES Cryptographic Provider";
    const auto cspName = utf16LeNullTerminated(kProvider);
    const auto headerSize = static_cast<std::uint32_t>(32u + cspName.size());
    Bytes out;
    appendU16(out, 4); appendU16(out, 2);
    appendU32(out, 0x24u); // fCryptoAPI | fAES
    appendU32(out, headerSize);
    appendU32(out, 0x24u); // EncryptionHeader.Flags
    appendU32(out, 0);     // SizeExtra
    appendU32(out, standardAlgId(keyBits));
    appendU32(out, 0x8004u); // SHA-1
    appendU32(out, keyBits);
    appendU32(out, 0x18u); // PROV_RSA_AES
    appendU32(out, 0);
    appendU32(out, 0);
    out.insert(out.end(), cspName.begin(), cspName.end());
    appendU32(out, static_cast<std::uint32_t>(salt.size()));
    out.insert(out.end(), salt.begin(), salt.end());
    out.insert(out.end(), encryptedVerifier.begin(), encryptedVerifier.end());
    appendU32(out, 20u);
    out.insert(out.end(), encryptedVerifierHash.begin(), encryptedVerifierHash.end());
    return out;
}

PackageEncryptionInfo inspectEncryptionInfoStream(const Bytes& encryptionInfo) {
    PackageEncryptionInfo out;
    out.encrypted = true;
    if (encryptionInfo.size() < 4) {
        out.format = PackageEncryptionFormat::Unsupported;
        return out;
    }
    out.versionMajor = getU16(encryptionInfo.data());
    out.versionMinor = getU16(encryptionInfo.data() + 2);
    if (out.versionMajor == 4 && out.versionMinor == 4) {
        out.format = PackageEncryptionFormat::Agile;
        try {
            if (encryptionInfo.size() < 8 || getU32(encryptionInfo.data() + 4) != 0x40) throw std::runtime_error("bad reserved");
            const std::string xml(encryptionInfo.begin() + 8, encryptionInfo.end());
            const auto keyData = elementByLocalName(xml, "keyData");
            const auto integrity = elementByLocalName(xml, "dataIntegrity");
            const auto keyEncryptors = agileKeyEncryptors(xml);
            out.keyEncryptorCount = keyEncryptors.size();
            const auto passwordUri = std::string("http://schemas.microsoft.com/office/2006/keyEncryptor/password");
            const auto certificateUri = std::string("http://schemas.microsoft.com/office/2006/keyEncryptor/certificate");
            const AgileKeyEncryptorXml* passwordKey = nullptr;
            for (const auto& item : keyEncryptors) {
                if (item.uri == passwordUri) {
                    ++out.passwordKeyEncryptorCount;
                    if (!passwordKey) passwordKey = &item;
                    continue;
                }
                if (item.uri != certificateUri) continue;
                PackageEncryptionCertificateKeyInfo certificate;
                try {
                    const auto encryptedKeyValue = base64Decode(attr(item.encryptedKeyElement, "encryptedKeyValue"));
                    const auto x509 = base64Decode(attr(item.encryptedKeyElement, "X509Certificate"));
                    const auto verifier = base64Decode(attr(item.encryptedKeyElement, "certVerifier"));
                    certificate.encryptedKeyBytes = encryptedKeyValue.size();
                    certificate.certVerifierBytes = verifier.size();
                    certificate.x509Certificate = x509;
                    certificate.validEncoding = !x509.empty() && !encryptedKeyValue.empty() && !verifier.empty();
                } catch (...) {
                    certificate.validEncoding = false;
                }
                out.certificateKeyEncryptors.push_back(std::move(certificate));
            }
            if (keyData.empty() || !passwordKey || passwordKey->encryptedKeyElement.empty())
                throw std::runtime_error("missing Agile password descriptor");
            out.cipherAlgorithm = attr(keyData, "cipherAlgorithm");
            out.cipherChaining = attr(keyData, "cipherChaining");
            out.keyBits = uintAttr(keyData, "keyBits");
            out.blockSize = uintAttr(keyData, "blockSize");
            out.hashSize = uintAttr(keyData, "hashSize");
            out.hashAlgorithm = parseOfficeHashName(attr(keyData, "hashAlgorithm"));
            out.spinCount = uintAttr(passwordKey->encryptedKeyElement, "spinCount");
            out.hasDataIntegrity = !integrity.empty();
            const auto parsed = parseAgileInfo(encryptionInfo);
            (void)parsed;
            // Certificate key-encryptors are inspected/preserved as metadata in
            // P1I. Password-based decryption remains the supported read path.
            out.supportedForRead = true;
            out.supportedForWrite = true;
        } catch (...) {
            out.supportedForRead = false;
            out.supportedForWrite = false;
        }
        return out;
    }
    if (out.versionMinor == 2 && (out.versionMajor == 2 || out.versionMajor == 3 || out.versionMajor == 4)) {
        out.format = PackageEncryptionFormat::Standard;
        out.keyEncryptorCount = 1;
        out.passwordKeyEncryptorCount = 1;
        out.cipherChaining = "ChainingModeECB";
        out.hashAlgorithm = PackageEncryptionHash::Sha1;
        out.hashSize = 20;
        out.blockSize = 16;
        out.spinCount = 50000;
        try {
            const auto parsed = parseStandardInfo(encryptionInfo);
            out.keyBits = static_cast<std::uint32_t>(parsed.keyBytes * 8u);
            out.cipherAlgorithm = "AES";
            out.supportedForRead = true;
            out.supportedForWrite = true;
        } catch (...) {
            out.supportedForRead = false;
            out.supportedForWrite = false;
        }
        return out;
    }
    out.format = PackageEncryptionFormat::Unsupported;
    return out;
}

} // namespace

std::vector<unsigned char> readBinaryFile(const std::filesystem::path& path) {
    std::ifstream stream(path, std::ios::binary);
    if (!stream) throw std::runtime_error("Unable to open file: " + path.string());
    stream.seekg(0, std::ios::end);
    const auto end = stream.tellg();
    if (end < 0) throw std::runtime_error("Unable to determine file size: " + path.string());
    Bytes bytes(static_cast<std::size_t>(end));
    stream.seekg(0, std::ios::beg);
    if (!bytes.empty()) stream.read(reinterpret_cast<char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
    if (!stream && !bytes.empty()) throw std::runtime_error("Unable to read file: " + path.string());
    return bytes;
}

void writeBinaryFile(const std::filesystem::path& path, const std::vector<unsigned char>& bytes) {
    atomicWriteFile(path, [&](const std::filesystem::path& temporary) {
        std::ofstream stream(temporary, std::ios::binary | std::ios::trunc);
        if (!stream) throw std::runtime_error("Unable to create file: " + temporary.string());
        if (!bytes.empty())
            stream.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
        stream.flush();
        if (!stream) throw std::runtime_error("Unable to write file: " + temporary.string());
        stream.close();
        if (!stream) throw std::runtime_error("Unable to close file: " + temporary.string());
    });
}

bool isEncryptedOfficeCompoundFile(const std::filesystem::path& path) noexcept {
    try {
        const auto bytes = readBinaryFile(path);
        return isCompoundFile(bytes) && compoundFileContainsStream(bytes, "EncryptionInfo") && compoundFileContainsStream(bytes, "EncryptedPackage");
    } catch (...) { return false; }
}

PackageEncryptionInfo inspectOfficeEncryption(const std::vector<unsigned char>& compoundFileBytes) {
    if (!isCompoundFile(compoundFileBytes) || !compoundFileContainsStream(compoundFileBytes, "EncryptionInfo") ||
        !compoundFileContainsStream(compoundFileBytes, "EncryptedPackage"))
        return {};
    const auto encryptionInfo = readCompoundFileStream(compoundFileBytes, "EncryptionInfo");
    return inspectEncryptionInfoStream(encryptionInfo);
}

PackageEncryptionInfo inspectOfficeEncryption(const std::filesystem::path& path) {
    return inspectOfficeEncryption(readBinaryFile(path));
}

std::vector<unsigned char> encryptAgileOfficePackage(const std::vector<unsigned char>& packageBytes,
                                                     const std::string& password,
                                                     const AgileEncryptionParameters& parameters) {
    if (parameters.spinCount > 10000000u) throw std::invalid_argument("Agile Encryption spinCount must be <= 10,000,000");
    if (parameters.keyBits != 128 && parameters.keyBits != 192 && parameters.keyBits != 256)
        throw std::invalid_argument("Agile Encryption keyBits must be 128, 192, or 256");
    const auto keyBytes = static_cast<std::size_t>(parameters.keyBits / 8u);
    const auto keyDataSalt = randomBytes(kSaltSize);
    const auto passwordSalt = randomBytes(kSaltSize);
    const auto verifier = randomBytes(kSaltSize);
    auto intermediateKey = randomBytes(keyBytes);
    WipeOnExit wipeIntermediate{&intermediateKey};

    const auto passwordIv = deriveIv(passwordSalt, parameters.hashAlgorithm, nullptr, 0, kBlockSize);
    auto verifierKey = derivePasswordKey(password, passwordSalt, parameters.spinCount, parameters.hashAlgorithm,
        kVerifierHashInputBlockKey.data(), kVerifierHashInputBlockKey.size(), keyBytes);
    auto verifierHashKey = derivePasswordKey(password, passwordSalt, parameters.spinCount, parameters.hashAlgorithm,
        kVerifierHashValueBlockKey.data(), kVerifierHashValueBlockKey.size(), keyBytes);
    auto encryptedKeyKey = derivePasswordKey(password, passwordSalt, parameters.spinCount, parameters.hashAlgorithm,
        kEncryptedKeyValueBlockKey.data(), kEncryptedKeyValueBlockKey.size(), keyBytes);
    WipeOnExit wipeVerifierKey{&verifierKey}, wipeVerifierHashKey{&verifierHashKey}, wipeEncryptedKeyKey{&encryptedKeyKey};

    const auto encryptedVerifierInput = aesCbc(zeroPad(verifier, kBlockSize), verifierKey, passwordIv, true);
    const auto encryptedVerifierHash = aesCbc(zeroPad(hashDigest(parameters.hashAlgorithm, verifier), kBlockSize), verifierHashKey, passwordIv, true);
    const auto encryptedKey = aesCbc(zeroPad(intermediateKey, kBlockSize), encryptedKeyKey, passwordIv, true);

    const auto encryptedPackage = encryptPackageSegments(packageBytes, intermediateKey, keyDataSalt,
        parameters.hashAlgorithm, kBlockSize);
    auto integrityKey = randomBytes(keyDataSalt.size());
    WipeOnExit wipeIntegrityKey{&integrityKey};
    const auto integrityValue = hmacDigest(parameters.hashAlgorithm, integrityKey, encryptedPackage);
    const auto integrityKeyIv = deriveIv(keyDataSalt, parameters.hashAlgorithm,
        kIntegrityKeyBlockKey.data(), kIntegrityKeyBlockKey.size(), kBlockSize);
    const auto integrityValueIv = deriveIv(keyDataSalt, parameters.hashAlgorithm,
        kIntegrityValueBlockKey.data(), kIntegrityValueBlockKey.size(), kBlockSize);
    const auto encryptedHmacKey = aesCbc(zeroPad(integrityKey, kBlockSize), intermediateKey, integrityKeyIv, true);
    const auto encryptedHmacValue = aesCbc(zeroPad(integrityValue, kBlockSize), intermediateKey, integrityValueIv, true);

    Bytes encryptionInfo;
    appendU16(encryptionInfo, 4);
    appendU16(encryptionInfo, 4);
    appendU32(encryptionInfo, 0x40);
    const auto xml = agileXml(parameters.spinCount, parameters.keyBits, parameters.hashAlgorithm,
        keyDataSalt, passwordSalt, encryptedVerifierInput, encryptedVerifierHash, encryptedKey,
        encryptedHmacKey, encryptedHmacValue);
    encryptionInfo.insert(encryptionInfo.end(), xml.begin(), xml.end());

    return buildRootCompoundFile({
        {"EncryptedPackage", encryptedPackage},
        {"EncryptionInfo", encryptionInfo},
    });
}

std::vector<unsigned char> encryptStandardOfficePackage(const std::vector<unsigned char>& packageBytes,
                                                        const std::string& password,
                                                        const StandardEncryptionParameters& parameters) {
    if (parameters.keyBits != 128 && parameters.keyBits != 192 && parameters.keyBits != 256)
        throw std::invalid_argument("Standard Encryption keyBits must be 128, 192, or 256");
    const auto keyBytes = static_cast<std::size_t>(parameters.keyBits / 8u);
    const auto salt = randomBytes(kSaltSize);
    const auto verifier = randomBytes(kSaltSize);
    auto key = deriveStandardKey(password, salt, keyBytes);
    WipeOnExit wipeKey{&key};
    const auto encryptedVerifier = aesEcb(verifier, key, true);
    const auto verifierHash = hashDigest(PackageEncryptionHash::Sha1, verifier);
    const auto encryptedVerifierHash = aesEcb(zeroPad(verifierHash, kBlockSize), key, true);
    const auto encryptionInfo = standardEncryptionInfo(salt, encryptedVerifier, encryptedVerifierHash, parameters.keyBits);

    Bytes encryptedPackage;
    appendU64(encryptedPackage, static_cast<std::uint64_t>(packageBytes.size()));
    const auto padded = zeroPad(packageBytes, kBlockSize);
    const auto cipher = aesEcb(padded, key, true);
    encryptedPackage.insert(encryptedPackage.end(), cipher.begin(), cipher.end());

    return buildRootCompoundFile({
        {"EncryptedPackage", encryptedPackage},
        {"EncryptionInfo", encryptionInfo},
    });
}

std::vector<unsigned char> decryptOfficePackage(const std::vector<unsigned char>& compoundFileBytes,
                                                const std::string& password,
                                                const OfficeDecryptionLimits& limits) {
    if (!isCompoundFile(compoundFileBytes)) throw std::runtime_error("Encrypted Office file is not a CFB compound document");
    const auto encryptionInfo = readCompoundFileStream(compoundFileBytes, "EncryptionInfo");
    const auto encryptedPackage = readCompoundFileStream(compoundFileBytes, "EncryptedPackage");
    if (limits.maxEncryptionInfoBytes != 0 && encryptionInfo.size() > limits.maxEncryptionInfoBytes)
        throw std::runtime_error("EncryptionInfo exceeds configured maxEncryptionInfoBytes limit");
    if (encryptionInfo.size() < 4) throw std::runtime_error("EncryptionInfo stream is truncated");
    const auto major = getU16(encryptionInfo.data());
    const auto minor = getU16(encryptionInfo.data() + 2);
    if (major == 4 && minor == 4) {
        const auto info = parseAgileInfo(encryptionInfo);
        const auto effectiveSpinLimit = limits.maxSpinCount == 0 ? 10000000u : limits.maxSpinCount;
        if (info.spinCount > effectiveSpinLimit)
            throw std::runtime_error("Agile Encryption spinCount exceeds configured maxEncryptionSpinCount limit");
        auto intermediateKey = decryptIntermediateKey(info, password);
        WipeOnExit wipeIntermediate{&intermediateKey};
        if (limits.requireAgileDataIntegrity && info.encryptedHmacKey.empty() && info.encryptedHmacValue.empty())
            throw std::runtime_error("Agile Encryption is rejected by policy because DataIntegrity is required");
        if (!info.encryptedHmacKey.empty() || !info.encryptedHmacValue.empty()) {
            if (info.encryptedHmacKey.empty() || info.encryptedHmacValue.empty())
                throw std::runtime_error("Agile Encryption dataIntegrity element is incomplete");
            const auto keyIv = deriveIv(info.keyData.salt, info.keyData.hashAlgorithm,
                kIntegrityKeyBlockKey.data(), kIntegrityKeyBlockKey.size(), info.keyData.blockSize);
            const auto valueIv = deriveIv(info.keyData.salt, info.keyData.hashAlgorithm,
                kIntegrityValueBlockKey.data(), kIntegrityValueBlockKey.size(), info.keyData.blockSize);
            auto integrityKey = aesCbc(info.encryptedHmacKey, intermediateKey, keyIv, false);
            WipeOnExit wipeIntegrityKey{&integrityKey};
            integrityKey.resize(info.keyData.salt.size());
            const auto expected = aesCbc(info.encryptedHmacValue, intermediateKey, valueIv, false);
            const auto actual = hmacDigest(info.keyData.hashAlgorithm, integrityKey, encryptedPackage);
            if (expected.size() < actual.size() || !constantTimeEqual(expected.data(), actual.data(), actual.size()))
                throw std::runtime_error("Encrypted Office package failed HMAC integrity verification");
        }
        return decryptPackageSegments(encryptedPackage, intermediateKey, info.keyData.salt,
            info.keyData.hashAlgorithm, info.keyData.blockSize, limits);
    }
    if (minor == 2 && (major == 2 || major == 3 || major == 4)) {
        if (!limits.allowStandardEncryption)
            throw std::runtime_error("Standard Encryption is disabled by LoadOptions policy");
        return decryptStandardPackage(encryptionInfo, encryptedPackage, password, limits);
    }
    throw std::runtime_error("Unsupported ECMA-376 encryption version");
}


} // namespace xlpp::internal
