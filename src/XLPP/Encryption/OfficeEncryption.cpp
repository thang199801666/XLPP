#include "OfficeEncryption.h"
#if defined(_WIN32) && !defined(NOMINMAX)
#define NOMINMAX
#endif
#include <algorithm>
#include <array>
#include <cctype>
#include <cstdint>
#include <cstring>
#include <fstream>
#include <limits>
#include <map>
#include <optional>
#include <random>
#include <stdexcept>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#if defined(_WIN32)
#include <windows.h>
#include <bcrypt.h>
#elif defined(__linux__)
#include <sys/random.h>
#elif defined(__APPLE__) || defined(__FreeBSD__)
#include <stdlib.h>
#endif
namespace xlpp::internal {
    namespace {
        using Bytes = std::vector<unsigned char>;
        constexpr std::uint32_t kFreeSect = 0xFFFFFFFFu;
        constexpr std::uint32_t kEndOfChain = 0xFFFFFFFEu;
        constexpr std::uint32_t kFatSect = 0xFFFFFFFDu;
        constexpr std::uint32_t kDifSect = 0xFFFFFFFCu;
        constexpr std::size_t kSectorSize = 512;
        constexpr std::size_t kMiniSectorSize = 64;
        constexpr std::size_t kMiniCutoff = 4096;
        std::uint16_t getU16(const unsigned char* p) {
            return static_cast<std::uint16_t>(p[0]) | (static_cast<std::uint16_t>(p[1]) << 8u);
        }
        std::uint32_t getU32(const unsigned char* p) {
            return static_cast<std::uint32_t>(p[0]) | (static_cast<std::uint32_t>(p[1]) << 8u) | (static_cast<std::uint32_t>(p[2]) << 16u) | (static_cast<std::uint32_t>(p[3]) << 24u);
        }
        std::uint64_t getU64(const unsigned char* p) {
            std::uint64_t value = 0;
            for (unsigned shift = 0; shift < 64; shift += 8) value |= static_cast<std::uint64_t>(p[shift / 8]) << shift;
            return value;
        }
        void putU16(Bytes& out, std::uint16_t value) {
            out.push_back(static_cast<unsigned char>(value));
            out.push_back(static_cast<unsigned char>(value >> 8u));
        }
        void putU32(Bytes& out, std::uint32_t value) {
            for (unsigned shift = 0; shift < 32; shift += 8) out.push_back(static_cast<unsigned char>(value >> shift));
        }
        void putU64(Bytes& out, std::uint64_t value) {
            for (unsigned shift = 0; shift < 64; shift += 8) out.push_back(static_cast<unsigned char>(value >> shift));
        }
        void overwriteU16(Bytes& out, std::size_t offset, std::uint16_t value) {
            out.at(offset) = static_cast<unsigned char>(value);
            out.at(offset + 1) = static_cast<unsigned char>(value >> 8u);
        }
        void overwriteU32(Bytes& out, std::size_t offset, std::uint32_t value) {
            for (unsigned shift = 0; shift < 32; shift += 8) out.at(offset + shift / 8) = static_cast<unsigned char>(value >> shift);
        }
        std::string lowerAscii(std::string value) {
            std::transform(value.begin(), value.end(), value.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
            return value;
        }
        Bytes randomBytes(std::size_t count) {
            Bytes out(count);
            if (out.empty()) return out;
            #if defined(_WIN32)
            if (out.size() <= static_cast<std::size_t>(std::numeric_limits<ULONG>::max()) && BCryptGenRandom(nullptr, reinterpret_cast<PUCHAR>(out.data()), static_cast<ULONG>(out.size()), BCRYPT_USE_SYSTEM_PREFERRED_RNG) == 0) return out;
            #elif defined(__linux__)
            std::size_t offset = 0;
            while (offset < out.size()) {
                const auto countRead = ::getrandom(out.data() + offset, out.size() - offset, 0);
                if (countRead <= 0) break;
                offset += static_cast<std::size_t>(countRead);
            }
            if (offset == out.size()) return out;
            #elif defined(__APPLE__) || defined(__FreeBSD__)
            arc4random_buf(out.data(), out.size());
            return out;
            #endif
            std::random_device source;
            for (auto& byte : out) byte = static_cast<unsigned char>(source());
            return out;
        }
        // ---------- SHA-512 / HMAC-SHA-512 ----------
        constexpr std::array<std::uint64_t, 80> kSha512K {
            {
                0x428a2f98d728ae22ULL,0x7137449123ef65cdULL,0xb5c0fbcfec4d3b2fULL,0xe9b5dba58189dbbcULL, 0x3956c25bf348b538ULL,0x59f111f1b605d019ULL,0x923f82a4af194f9bULL,0xab1c5ed5da6d8118ULL, 0xd807aa98a3030242ULL,0x12835b0145706fbeULL,0x243185be4ee4b28cULL,0x550c7dc3d5ffb4e2ULL, 0x72be5d74f27b896fULL,0x80deb1fe3b1696b1ULL,0x9bdc06a725c71235ULL,0xc19bf174cf692694ULL, 0xe49b69c19ef14ad2ULL,0xefbe4786384f25e3ULL,0x0fc19dc68b8cd5b5ULL,0x240ca1cc77ac9c65ULL, 0x2de92c6f592b0275ULL,0x4a7484aa6ea6e483ULL,0x5cb0a9dcbd41fbd4ULL,0x76f988da831153b5ULL, 0x983e5152ee66dfabULL,0xa831c66d2db43210ULL,0xb00327c898fb213fULL,0xbf597fc7beef0ee4ULL, 0xc6e00bf33da88fc2ULL,0xd5a79147930aa725ULL,0x06ca6351e003826fULL,0x142929670a0e6e70ULL, 0x27b70a8546d22ffcULL,0x2e1b21385c26c926ULL,0x4d2c6dfc5ac42aedULL,0x53380d139d95b3dfULL, 0x650a73548baf63deULL,0x766a0abb3c77b2a8ULL,0x81c2c92e47edaee6ULL,0x92722c851482353bULL, 0xa2bfe8a14cf10364ULL,0xa81a664bbc423001ULL,0xc24b8b70d0f89791ULL,0xc76c51a30654be30ULL, 0xd192e819d6ef5218ULL,0xd69906245565a910ULL,0xf40e35855771202aULL,0x106aa07032bbd1b8ULL, 0x19a4c116b8d2d0c8ULL,0x1e376c085141ab53ULL,0x2748774cdf8eeb99ULL,0x34b0bcb5e19b48a8ULL, 0x391c0cb3c5c95a63ULL,0x4ed8aa4ae3418acbULL,0x5b9cca4f7763e373ULL,0x682e6ff3d6b2b8a3ULL, 0x748f82ee5defb2fcULL,0x78a5636f43172f60ULL,0x84c87814a1f0ab72ULL,0x8cc702081a6439ecULL, 0x90befffa23631e28ULL,0xa4506cebde82bde9ULL,0xbef9a3f7b2c67915ULL,0xc67178f2e372532bULL, 0xca273eceea26619cULL,0xd186b8c721c0c207ULL,0xeada7dd6cde0eb1eULL,0xf57d4f7fee6ed178ULL, 0x06f067aa72176fbaULL,0x0a637dc5a2c898a6ULL,0x113f9804bef90daeULL,0x1b710b35131c471bULL, 0x28db77f523047d84ULL,0x32caab7b40c72493ULL,0x3c9ebe0a15c9bebcULL,0x431d67c49c100d4cULL, 0x4cc5d4becb3e42b6ULL,0x597f299cfc657e2aULL,0x5fcb6fab3ad6faecULL,0x6c44198c4a475817ULL
            }
        };
        std::uint64_t rotr64(std::uint64_t v, unsigned n) {
            return (v >> n) | (v << (64 - n));
        }
        Bytes sha512(const Bytes& input) {
            Bytes data = input;
            const std::uint64_t bitLow = static_cast<std::uint64_t>(data.size()) * 8ULL;
            const std::uint64_t bitHigh = 0;
            data.push_back(0x80);
            while (data.size() % 128 != 112) data.push_back(0);
            for (int i = 7; i >= 0; --i) data.push_back(static_cast<unsigned char>(bitHigh >> (i * 8)));
            for (int i = 7; i >= 0; --i) data.push_back(static_cast<unsigned char>(bitLow >> (i * 8)));
            std::array<std::uint64_t,8> h {
                {
                    0x6a09e667f3bcc908ULL,0xbb67ae8584caa73bULL,0x3c6ef372fe94f82bULL,0xa54ff53a5f1d36f1ULL, 0x510e527fade682d1ULL,0x9b05688c2b3e6c1fULL,0x1f83d9abfb41bd6bULL,0x5be0cd19137e2179ULL
                }
            };
            for (std::size_t offset = 0; offset < data.size(); offset += 128) {
                std::array<std::uint64_t,80> w {
                };
                for (std::size_t i = 0; i < 16; ++i) {
                    std::uint64_t value = 0;
                    for (std::size_t j = 0; j < 8; ++j) value = (value << 8) | data[offset + i * 8 + j];
                    w[i] = value;
                }
                for (std::size_t i = 16; i < 80; ++i) {
                    const auto s0 = rotr64(w[i-15],1) ^ rotr64(w[i-15],8) ^ (w[i-15] >> 7);
                    const auto s1 = rotr64(w[i-2],19) ^ rotr64(w[i-2],61) ^ (w[i-2] >> 6);
                    w[i] = w[i-16] + s0 + w[i-7] + s1;
                }
                auto a=h[0],b=h[1],c=h[2],d=h[3],e=h[4],f=h[5],g=h[6],hh=h[7];
                for (std::size_t i = 0; i < 80; ++i) {
                    const auto s1 = rotr64(e,14) ^ rotr64(e,18) ^ rotr64(e,41);
                    const auto ch = (e & f) ^ ((~e) & g);
                    const auto t1 = hh + s1 + ch + kSha512K[i] + w[i];
                    const auto s0 = rotr64(a,28) ^ rotr64(a,34) ^ rotr64(a,39);
                    const auto maj = (a & b) ^ (a & c) ^ (b & c);
                    const auto t2 = s0 + maj;
                    hh=g;
                    g=f;
                    f=e;
                    e=d+t1;
                    d=c;
                    c=b;
                    b=a;
                    a=t1+t2;
                }
                h[0]+=a;
                h[1]+=b;
                h[2]+=c;
                h[3]+=d;
                h[4]+=e;
                h[5]+=f;
                h[6]+=g;
                h[7]+=hh;
            }
            Bytes out;
            out.reserve(64);
            for (auto value : h) for (int i = 7; i >= 0; --i) out.push_back(static_cast<unsigned char>(value >> (i * 8)));
            return out;
        }
        Bytes hmacSha512(Bytes key, const Bytes& message) {
            constexpr std::size_t block = 128;
            if (key.size() > block) key = sha512(key);
            key.resize(block, 0);
            Bytes inner(block), outer(block);
            for (std::size_t i = 0; i < block; ++i) {
                inner[i] = key[i] ^ 0x36;
                outer[i] = key[i] ^ 0x5c;
            }
            inner.insert(inner.end(), message.begin(), message.end());
            const auto innerHash = sha512(inner);
            outer.insert(outer.end(), innerHash.begin(), innerHash.end());
            return sha512(outer);
        }
        // ---------- SHA-1 (ECMA-376 Standard Encryption) ----------
        std::uint32_t rotl32(std::uint32_t v, unsigned n) {
            return (v << n) | (v >> (32 - n));
        }
        Bytes sha1(const Bytes& input) {
            Bytes data = input;
            const std::uint64_t bitLength = static_cast<std::uint64_t>(data.size()) * 8ULL;
            data.push_back(0x80);
            while (data.size() % 64 != 56) data.push_back(0);
            for (int i = 7; i >= 0; --i) data.push_back(static_cast<unsigned char>(bitLength >> (i * 8)));
            std::uint32_t h0=0x67452301u,h1=0xefcdab89u,h2=0x98badcfeu,h3=0x10325476u,h4=0xc3d2e1f0u;
            for (std::size_t offset=0; offset<data.size(); offset+=64) {
                std::array<std::uint32_t,80> w {
                };
                for (std::size_t i=0;i<16;++i) {
                    auto q=data.data()+offset+i*4;
                    w[i]=(static_cast<std::uint32_t>(q[0])<<24)|(static_cast<std::uint32_t>(q[1])<<16)| (static_cast<std::uint32_t>(q[2])<<8)|static_cast<std::uint32_t>(q[3]);
                }
                for (std::size_t i=16;i<80;++i) w[i]=rotl32(w[i-3]^w[i-8]^w[i-14]^w[i-16],1);
                auto a=h0,b=h1,c=h2,d=h3,e=h4;
                for (std::size_t i=0;i<80;++i) {
                    std::uint32_t f,k;
                    if(i<20) {
                        f=(b&c)|((~b)&d);
                        k=0x5a827999u;
                    } else if(i<40) {
                        f=b^c^d;
                        k=0x6ed9eba1u;
                    } else if(i<60) {
                        f=(b&c)|(b&d)|(c&d);
                        k=0x8f1bbcdcu;
                    } else {
                        f=b^c^d;
                        k=0xca62c1d6u;
                    }
                    auto temp=rotl32(a,5)+f+e+k+w[i];
                    e=d;
                    d=c;
                    c=rotl32(b,30);
                    b=a;
                    a=temp;
                }
                h0+=a;
                h1+=b;
                h2+=c;
                h3+=d;
                h4+=e;
            }
            Bytes out;
            out.reserve(20);
            for(auto value:{h0,h1,h2,h3,h4}) for(int i=3;i>=0;--i) out.push_back(static_cast<unsigned char>(value>>(i*8)));
            return out;
        }
        // ---------- AES (128/192/256) ----------
        constexpr std::array<unsigned char,256> kSbox {
            {
                0x63,0x7c,0x77,0x7b,0xf2,0x6b,0x6f,0xc5,0x30,0x01,0x67,0x2b,0xfe,0xd7,0xab,0x76, 0xca,0x82,0xc9,0x7d,0xfa,0x59,0x47,0xf0,0xad,0xd4,0xa2,0xaf,0x9c,0xa4,0x72,0xc0, 0xb7,0xfd,0x93,0x26,0x36,0x3f,0xf7,0xcc,0x34,0xa5,0xe5,0xf1,0x71,0xd8,0x31,0x15, 0x04,0xc7,0x23,0xc3,0x18,0x96,0x05,0x9a,0x07,0x12,0x80,0xe2,0xeb,0x27,0xb2,0x75, 0x09,0x83,0x2c,0x1a,0x1b,0x6e,0x5a,0xa0,0x52,0x3b,0xd6,0xb3,0x29,0xe3,0x2f,0x84, 0x53,0xd1,0x00,0xed,0x20,0xfc,0xb1,0x5b,0x6a,0xcb,0xbe,0x39,0x4a,0x4c,0x58,0xcf, 0xd0,0xef,0xaa,0xfb,0x43,0x4d,0x33,0x85,0x45,0xf9,0x02,0x7f,0x50,0x3c,0x9f,0xa8, 0x51,0xa3,0x40,0x8f,0x92,0x9d,0x38,0xf5,0xbc,0xb6,0xda,0x21,0x10,0xff,0xf3,0xd2, 0xcd,0x0c,0x13,0xec,0x5f,0x97,0x44,0x17,0xc4,0xa7,0x7e,0x3d,0x64,0x5d,0x19,0x73, 0x60,0x81,0x4f,0xdc,0x22,0x2a,0x90,0x88,0x46,0xee,0xb8,0x14,0xde,0x5e,0x0b,0xdb, 0xe0,0x32,0x3a,0x0a,0x49,0x06,0x24,0x5c,0xc2,0xd3,0xac,0x62,0x91,0x95,0xe4,0x79, 0xe7,0xc8,0x37,0x6d,0x8d,0xd5,0x4e,0xa9,0x6c,0x56,0xf4,0xea,0x65,0x7a,0xae,0x08, 0xba,0x78,0x25,0x2e,0x1c,0xa6,0xb4,0xc6,0xe8,0xdd,0x74,0x1f,0x4b,0xbd,0x8b,0x8a, 0x70,0x3e,0xb5,0x66,0x48,0x03,0xf6,0x0e,0x61,0x35,0x57,0xb9,0x86,0xc1,0x1d,0x9e, 0xe1,0xf8,0x98,0x11,0x69,0xd9,0x8e,0x94,0x9b,0x1e,0x87,0xe9,0xce,0x55,0x28,0xdf, 0x8c,0xa1,0x89,0x0d,0xbf,0xe6,0x42,0x68,0x41,0x99,0x2d,0x0f,0xb0,0x54,0xbb,0x16
            }
        };
        constexpr std::array<unsigned char,256> kInvSbox {
            {
                0x52,0x09,0x6a,0xd5,0x30,0x36,0xa5,0x38,0xbf,0x40,0xa3,0x9e,0x81,0xf3,0xd7,0xfb, 0x7c,0xe3,0x39,0x82,0x9b,0x2f,0xff,0x87,0x34,0x8e,0x43,0x44,0xc4,0xde,0xe9,0xcb, 0x54,0x7b,0x94,0x32,0xa6,0xc2,0x23,0x3d,0xee,0x4c,0x95,0x0b,0x42,0xfa,0xc3,0x4e, 0x08,0x2e,0xa1,0x66,0x28,0xd9,0x24,0xb2,0x76,0x5b,0xa2,0x49,0x6d,0x8b,0xd1,0x25, 0x72,0xf8,0xf6,0x64,0x86,0x68,0x98,0x16,0xd4,0xa4,0x5c,0xcc,0x5d,0x65,0xb6,0x92, 0x6c,0x70,0x48,0x50,0xfd,0xed,0xb9,0xda,0x5e,0x15,0x46,0x57,0xa7,0x8d,0x9d,0x84, 0x90,0xd8,0xab,0x00,0x8c,0xbc,0xd3,0x0a,0xf7,0xe4,0x58,0x05,0xb8,0xb3,0x45,0x06, 0xd0,0x2c,0x1e,0x8f,0xca,0x3f,0x0f,0x02,0xc1,0xaf,0xbd,0x03,0x01,0x13,0x8a,0x6b, 0x3a,0x91,0x11,0x41,0x4f,0x67,0xdc,0xea,0x97,0xf2,0xcf,0xce,0xf0,0xb4,0xe6,0x73, 0x96,0xac,0x74,0x22,0xe7,0xad,0x35,0x85,0xe2,0xf9,0x37,0xe8,0x1c,0x75,0xdf,0x6e, 0x47,0xf1,0x1a,0x71,0x1d,0x29,0xc5,0x89,0x6f,0xb7,0x62,0x0e,0xaa,0x18,0xbe,0x1b, 0xfc,0x56,0x3e,0x4b,0xc6,0xd2,0x79,0x20,0x9a,0xdb,0xc0,0xfe,0x78,0xcd,0x5a,0xf4, 0x1f,0xdd,0xa8,0x33,0x88,0x07,0xc7,0x31,0xb1,0x12,0x10,0x59,0x27,0x80,0xec,0x5f, 0x60,0x51,0x7f,0xa9,0x19,0xb5,0x4a,0x0d,0x2d,0xe5,0x7a,0x9f,0x93,0xc9,0x9c,0xef, 0xa0,0xe0,0x3b,0x4d,0xae,0x2a,0xf5,0xb0,0xc8,0xeb,0xbb,0x3c,0x83,0x53,0x99,0x61, 0x17,0x2b,0x04,0x7e,0xba,0x77,0xd6,0x26,0xe1,0x69,0x14,0x63,0x55,0x21,0x0c,0x7d
            }
        };
        constexpr std::array<unsigned char,15> kRcon {
            {
                0x00,0x01,0x02,0x04,0x08,0x10,0x20,0x40,0x80,0x1b,0x36,0x6c,0xd8,0xab,0x4d
            }
        };
        unsigned char gmul(unsigned char a, unsigned char b) {
            unsigned char p = 0;
            for (int i = 0; i < 8; ++i) {
                if (b & 1) p ^= a;
                const bool high = (a & 0x80) != 0;
                a <<= 1;
                if (high) a ^= 0x1b;
                b >>= 1;
            }
            return p;
        }
        std::array<unsigned char,256> makeGaloisTable(unsigned char factor) {
            std::array<unsigned char,256> table {
            };
            for (std::size_t i=0;i<table.size();++i) table[i]=gmul(static_cast<unsigned char>(i),factor);
            return table;
        }
        const auto kMul2=makeGaloisTable(2),kMul3=makeGaloisTable(3),kMul9=makeGaloisTable(9), kMul11=makeGaloisTable(11),kMul13=makeGaloisTable(13),kMul14=makeGaloisTable(14);
        struct Aes {
            Bytes round;
            int rounds {
            };
            explicit Aes(const Bytes& key) {
                if (key.size()!=16 && key.size()!=24 && key.size()!=32) throw std::invalid_argument("AES key must be 16, 24 or 32 bytes");
                const std::size_t nk=key.size()/4;
                rounds=static_cast<int>(nk+6);
                round.resize(16*static_cast<std::size_t>(rounds+1));
                std::copy(key.begin(),key.end(),round.begin());
                std::size_t bytes=key.size();
                unsigned rcon=1;
                std::array<unsigned char,4> temp {
                };
                while(bytes<round.size()) {
                    for(std::size_t i=0;i<4;++i)temp[i]=round[bytes-4+i];
                    if(bytes%key.size()==0) {
                        const auto t=temp[0];
                        temp[0]=kSbox[temp[1]];
                        temp[1]=kSbox[temp[2]];
                        temp[2]=kSbox[temp[3]];
                        temp[3]=kSbox[t];
                        temp[0]^=kRcon[rcon++];
                    } else if(nk>6 && bytes%key.size()==16) {
                        for(auto&v:temp)v=kSbox[v];
                    }
                    for(std::size_t i=0;i<4 && bytes<round.size();++i) {
                        round[bytes]=round[bytes-key.size()]^temp[i];
                        ++bytes;
                    }
                }
            }
            void addRoundKey(std::array<unsigned char,16>& state,int r)const {
                const auto base = static_cast<std::size_t>(r) * 16u;
                for(std::size_t i=0;i<16;++i) state[i] ^= round[base + i];
            }
            static void sub(std::array<unsigned char,16>&s) {
                for(auto&v:s)v=kSbox[v];
            }
            static void invSub(std::array<unsigned char,16>&s) {
                for(auto&v:s)v=kInvSbox[v];
            }
            static void shift(std::array<unsigned char,16>&s) {
                auto t=s;
                for(std::size_t r=0;r<4;++r)for(std::size_t c=0;c<4;++c)s[r+4*c]=t[r+4*((c+r)%4)];
            }
            static void invShift(std::array<unsigned char,16>&s) {
                auto t=s;
                for(std::size_t r=0;r<4;++r)for(std::size_t c=0;c<4;++c)s[r+4*c]=t[r+4*((c+4-r)%4)];
            }
            static void mix(std::array<unsigned char,16>&s) {
                for(std::size_t c=0;c<4;++c) {
                    const std::size_t i=4*c;
                    auto a=s[i],b=s[i+1],d=s[i+2],e=s[i+3];
                    s[i]=kMul2[a]^kMul3[b]^d^e;
                    s[i+1]=a^kMul2[b]^kMul3[d]^e;
                    s[i+2]=a^b^kMul2[d]^kMul3[e];
                    s[i+3]=kMul3[a]^b^d^kMul2[e];
                }
            }
            static void invMix(std::array<unsigned char,16>&s) {
                for(std::size_t c=0;c<4;++c) {
                    const std::size_t i=4*c;
                    auto a=s[i],b=s[i+1],d=s[i+2],e=s[i+3];
                    s[i]=kMul14[a]^kMul11[b]^kMul13[d]^kMul9[e];
                    s[i+1]=kMul9[a]^kMul14[b]^kMul11[d]^kMul13[e];
                    s[i+2]=kMul13[a]^kMul9[b]^kMul14[d]^kMul11[e];
                    s[i+3]=kMul11[a]^kMul13[b]^kMul9[d]^kMul14[e];
                }
            }
            std::array<unsigned char,16> encrypt(std::array<unsigned char,16>s)const {
                addRoundKey(s,0);
                for(int r=1;r<rounds;++r) {
                    sub(s);
                    shift(s);
                    mix(s);
                    addRoundKey(s,r);
                }
                sub(s);
                shift(s);
                addRoundKey(s,rounds);
                return s;
            }
            std::array<unsigned char,16> decrypt(std::array<unsigned char,16>s)const {
                addRoundKey(s,rounds);
                for(int r=rounds-1;r>0;--r) {
                    invShift(s);
                    invSub(s);
                    addRoundKey(s,r);
                    invMix(s);
                }
                invShift(s);
                invSub(s);
                addRoundKey(s,0);
                return s;
            }
        };
        Bytes aesCbcEncrypt(const Bytes& plain,const Bytes& key,const Bytes&iv) {
            if(plain.size()%16||iv.size()!=16)throw std::invalid_argument("AES-CBC requires block-aligned input and 16-byte IV");
            Aes aes(key);
            Bytes out(plain.size());
            std::array<unsigned char,16>prev {
            };
            std::copy(iv.begin(),iv.end(),prev.begin());
            for(std::size_t offset=0;offset<plain.size();offset+=16) {
                std::array<unsigned char,16>b {
                };
                for(std::size_t i=0;i<16;++i)b[i]=plain[offset+i]^prev[i];
                b=aes.encrypt(b);
                std::copy(b.begin(),b.end(),out.begin()+static_cast<std::ptrdiff_t>(offset));
                prev=b;
            }
            return out;
        }
        Bytes aesCbcDecrypt(const Bytes& cipher,const Bytes&key,const Bytes&iv) {
            if(cipher.size()%16||iv.size()!=16)throw std::invalid_argument("AES-CBC requires block-aligned input and 16-byte IV");
            Aes aes(key);
            Bytes out(cipher.size());
            std::array<unsigned char,16>prev {
            };
            std::copy(iv.begin(),iv.end(),prev.begin());
            for(std::size_t offset=0;offset<cipher.size();offset+=16) {
                std::array<unsigned char,16>c {
                },
                b {
                };
                std::copy_n(cipher.begin()+static_cast<std::ptrdiff_t>(offset),16,c.begin());
                b=aes.decrypt(c);
                for(std::size_t i=0;i<16;++i)out[offset+i]=b[i]^prev[i];
                prev=c;
            }
            return out;
        }
        Bytes aesEcbEncrypt(const Bytes& plain,const Bytes&key) {
            if(plain.size()%16)throw std::invalid_argument("AES-ECB requires block-aligned input");
            Aes aes(key);
            Bytes out(plain.size());
            for(std::size_t offset=0;offset<plain.size();offset+=16) {
                std::array<unsigned char,16>b {
                };
                std::copy_n(plain.begin()+static_cast<std::ptrdiff_t>(offset),16,b.begin());
                b=aes.encrypt(b);
                std::copy(b.begin(),b.end(),out.begin()+static_cast<std::ptrdiff_t>(offset));
            }
            return out;
        }
        Bytes aesEcbDecrypt(const Bytes& cipher,const Bytes&key) {
            if(cipher.size()%16)throw std::invalid_argument("AES-ECB requires block-aligned input");
            Aes aes(key);
            Bytes out(cipher.size());
            for(std::size_t offset=0;offset<cipher.size();offset+=16) {
                std::array<unsigned char,16>b {
                };
                std::copy_n(cipher.begin()+static_cast<std::ptrdiff_t>(offset),16,b.begin());
                b=aes.decrypt(b);
                std::copy(b.begin(),b.end(),out.begin()+static_cast<std::ptrdiff_t>(offset));
            }
            return out;
        }
        Bytes zeroPad(Bytes value, std::size_t block=16) {
            if (value.size()%block) value.resize(((value.size()+block-1)/block)*block,0);
            return value;
        }
        // ---------- Base64 ----------
        const char* kBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
        std::string base64Encode(const Bytes& input) {
            std::string out;
            out.reserve((input.size()+2)/3*4);
            for(std::size_t i=0;i<input.size();i+=3) {
                std::uint32_t v=static_cast<std::uint32_t>(input[i])<<16;
                const bool b1=i+1<input.size(),b2=i+2<input.size();
                if(b1)v|=static_cast<std::uint32_t>(input[i+1])<<8;
                if(b2)v|=input[i+2];
                out.push_back(kBase64[(v>>18)&63]);
                out.push_back(kBase64[(v>>12)&63]);
                out.push_back(b1?kBase64[(v>>6)&63]:'=');
                out.push_back(b2?kBase64[v&63]:'=');
            }
            return out;
        }
        Bytes base64Decode(std::string_view input) {
            std::array<int,256> table {
            };
            table.fill(-1);
            for(int i=0;i<64;++i)table[static_cast<unsigned char>(kBase64[i])]=i;
            Bytes out;
            int val=0,bits=-8;
            for(char raw:input) {
                const auto c = static_cast<unsigned char>(raw);
                if(std::isspace(c))continue;
                if(c=='=')break;
                const int d=table[c];
                if(d<0)throw std::runtime_error("Invalid base64 in EncryptionInfo");
                val=(val<<6)+d;
                bits+=6;
                if(bits>=0) {
                    out.push_back(static_cast<unsigned char>((val>>bits)&0xff));
                    bits-=8;
                }
            }
            return out;
        }
        // ---------- UTF-8 -> UTF-16LE password ----------
        Bytes utf16LePassword(std::string_view input) {
            Bytes out;
            std::size_t i=0;
            auto put=[&](std::uint16_t u) {
                out.push_back(static_cast<unsigned char>(u));
                out.push_back(static_cast<unsigned char>(u>>8));
            };
            while(i<input.size()) {
                unsigned char c=static_cast<unsigned char>(input[i++]);
                std::uint32_t cp=0;
                unsigned extra=0;
                if(c<0x80)cp=c;
                else if((c&0xe0)==0xc0) {
                    cp=c&0x1f;
                    extra=1;
                } else if((c&0xf0)==0xe0) {
                    cp=c&0x0f;
                    extra=2;
                } else if((c&0xf8)==0xf0) {
                    cp=c&0x07;
                    extra=3;
                } else throw std::invalid_argument("Password is not valid UTF-8");
                for(unsigned n=0;n<extra;++n) {
                    if(i>=input.size())throw std::invalid_argument("Password is not valid UTF-8");
                    unsigned char d=static_cast<unsigned char>(input[i++]);
                    if((d&0xc0)!=0x80)throw std::invalid_argument("Password is not valid UTF-8");
                    cp=(cp<<6)|(d&0x3f);
                }
                if(cp<=0xffff) {
                    if(cp>=0xd800&&cp<=0xdfff)throw std::invalid_argument("Password contains invalid Unicode scalar");
                    put(static_cast<std::uint16_t>(cp));
                } else if(cp<=0x10ffff) {
                    cp-=0x10000;
                    put(static_cast<std::uint16_t>(0xd800+(cp>>10)));
                    put(static_cast<std::uint16_t>(0xdc00+(cp&0x3ff)));
                } else throw std::invalid_argument("Password contains invalid Unicode scalar");
            }
            return out;
        }
        Bytes append(const Bytes& a, const Bytes& b) {
            Bytes out=a;
            out.insert(out.end(),b.begin(),b.end());
            return out;
        }
        Bytes little32(std::uint32_t v) {
            Bytes out;
            putU32(out,v);
            return out;
        }
        Bytes deriveSpinHash(const Bytes& salt, const std::string& password, std::uint32_t spinCount) {
            auto h=sha512(append(salt,utf16LePassword(password)));
            for(std::uint32_t i=0;i<spinCount;++i) h=sha512(append(little32(i),h));
            return h;
        }
        Bytes deriveKey(const Bytes& spinHash, const Bytes& blockKey, std::size_t keyBytes=32) {
            auto key=sha512(append(spinHash,blockKey));
            if(key.size()<keyBytes)key.resize(keyBytes,0x36);
            else key.resize(keyBytes);
            return key;
        }
        Bytes deriveIv(const Bytes& salt, const std::optional<Bytes>& blockKey) {
            Bytes iv=blockKey?sha512(append(salt,*blockKey)):salt;
            if(iv.size()<16)iv.resize(16,0x36);
            else iv.resize(16);
            return iv;
        }
        constexpr std::array<unsigned char,8> kVerifierInputBlock {
            {
                0xfe,0xa7,0xd2,0x76,0x3b,0x4b,0x9e,0x79
            }
        };
        constexpr std::array<unsigned char,8> kVerifierHashBlock {
            {
                0xd7,0xaa,0x0f,0x6d,0x30,0x61,0x34,0x4e
            }
        };
        constexpr std::array<unsigned char,8> kEncryptedKeyBlock {
            {
                0x14,0x6e,0x0b,0xe7,0xab,0xac,0xd0,0xd6
            }
        };
        constexpr std::array<unsigned char,8> kHmacKeyBlock {
            {
                0x5f,0xb2,0xad,0x01,0x0c,0xb9,0xe1,0xf6
            }
        };
        constexpr std::array<unsigned char,8> kHmacValueBlock {
            {
                0xa0,0x67,0x7f,0x02,0xb2,0x2c,0x84,0x33
            }
        };
        Bytes asBytes(const std::array<unsigned char,8>& v) {
            return Bytes(v.begin(),v.end());
        }
        bool constantTimeEqual(const Bytes& lhs, const Bytes& rhs) noexcept {
            const std::size_t count = std::max(lhs.size(), rhs.size());
            std::size_t difference = lhs.size() ^ rhs.size();
            for (std::size_t i = 0; i < count; ++i) {
                const unsigned a = i < lhs.size() ? lhs[i] : 0u;
                const unsigned b = i < rhs.size() ? rhs[i] : 0u;
                difference |= a ^ b;
            }
            return difference == 0;
        }
        // ---------- CFB ----------
        struct CfbNode {
            std::string name;
            std::uint8_t type {
                2
            };
            int parent {
                -1
            };
            Bytes data;
            std::uint32_t start {
                kEndOfChain
            };
            std::uint64_t size {
                0
            };
            std::uint32_t left {
                kFreeSect
            },
            right {
                kFreeSect
            },
            child {
                kFreeSect
            };
            std::uint8_t color {
                1
            };
        };
        bool nodeLess(const CfbNode& a,const CfbNode& b) {
            auto l=lowerAscii(a.name),r=lowerAscii(b.name);
            if(l.size()!=r.size())return l.size()<r.size();
            return l<r;
        }
        void assignChildTree(std::vector<CfbNode>& nodes,std::uint32_t storageId) {
            std::vector<std::uint32_t> children;
            for(std::uint32_t i=1;i<nodes.size();++i)if(nodes[i].parent==static_cast<int>(storageId))children.push_back(i);
            for(auto id:children) {
                nodes[id].left=nodes[id].right=kFreeSect;
                nodes[id].color=0;
            }
            if(children.empty()) {
                nodes[storageId].child=kFreeSect;
                return;
            }
            std::vector<std::uint32_t> parent(nodes.size(),kFreeSect);
            std::uint32_t root=kFreeSect;
            auto rotateLeft=[&](std::uint32_t x) {
                auto y=nodes[x].right;
                nodes[x].right=nodes[y].left;
                if(nodes[y].left!=kFreeSect)parent[nodes[y].left]=x;
                parent[y]=parent[x];
                if(parent[x]==kFreeSect)root=y;
                else if(x==nodes[parent[x]].left)nodes[parent[x]].left=y;
                else nodes[parent[x]].right=y;
                nodes[y].left=x;
                parent[x]=y;
            };
            auto rotateRight=[&](std::uint32_t y) {
                auto x=nodes[y].left;
                nodes[y].left=nodes[x].right;
                if(nodes[x].right!=kFreeSect)parent[nodes[x].right]=y;
                parent[x]=parent[y];
                if(parent[y]==kFreeSect)root=x;
                else if(y==nodes[parent[y]].right)nodes[parent[y]].right=x;
                else nodes[parent[y]].left=x;
                nodes[x].right=y;
                parent[y]=x;
            };
            for(auto z0:children) {
                std::uint32_t p=kFreeSect,c=root;
                while(c!=kFreeSect) {
                    p=c;
                    c=nodeLess(nodes[z0],nodes[c])?nodes[c].left:nodes[c].right;
                }
                parent[z0]=p;
                if(p==kFreeSect)root=z0;
                else if(nodeLess(nodes[z0],nodes[p]))nodes[p].left=z0;
                else nodes[p].right=z0;
                auto z=z0;
                while(z!=root&&nodes[parent[z]].color==0) {
                    auto pp=parent[z],gp=parent[pp];
                    if(pp==nodes[gp].left) {
                        auto y=nodes[gp].right;
                        if(y!=kFreeSect&&nodes[y].color==0) {
                            nodes[pp].color=1;
                            nodes[y].color=1;
                            nodes[gp].color=0;
                            z=gp;
                        } else {
                            if(z==nodes[pp].right) {
                                z=pp;
                                rotateLeft(z);
                                pp=parent[z];
                                gp=parent[pp];
                            }
                            nodes[pp].color=1;
                            nodes[gp].color=0;
                            rotateRight(gp);
                        }
                    } else {
                        auto y=nodes[gp].left;
                        if(y!=kFreeSect&&nodes[y].color==0) {
                            nodes[pp].color=1;
                            nodes[y].color=1;
                            nodes[gp].color=0;
                            z=gp;
                        } else {
                            if(z==nodes[pp].left) {
                                z=pp;
                                rotateRight(z);
                                pp=parent[z];
                                gp=parent[pp];
                            }
                            nodes[pp].color=1;
                            nodes[gp].color=0;
                            rotateLeft(gp);
                        }
                    }
                }
                nodes[root].color=1;
            }
            nodes[storageId].child=root;
        }
        Bytes utf16LeName(std::string_view ascii,bool nullTerm=true) {
            Bytes out;
            for(char raw:ascii) {
                const auto c = static_cast<unsigned char>(raw);
                out.push_back(c);
                out.push_back(0);
            }
            if(nullTerm) {
                out.push_back(0);
                out.push_back(0);
            }
            return out;
        }
        std::array<unsigned char,128> directoryEntry(const CfbNode& node) {
            std::array<unsigned char,128>out {
            };
            const auto wide=utf16LeName(node.name);
            if(wide.size()>64)throw std::invalid_argument("CFB directory name too long: "+node.name);
            std::copy(wide.begin(),wide.end(),out.begin());
            out[64]=static_cast<unsigned char>(wide.size());
            out[65]=static_cast<unsigned char>(wide.size()>>8);
            out[66]=node.type;
            out[67]=node.color;
            auto w32=[&](std::size_t o,std::uint32_t v) {
                for(unsigned s=0;s<32;s+=8)out[o+s/8]=static_cast<unsigned char>(v>>s);
            };
            auto w64=[&](std::size_t o,std::uint64_t v) {
                for(unsigned s=0;s<64;s+=8)out[o+s/8]=static_cast<unsigned char>(v>>s);
            };
            w32(68,node.left);
            w32(72,node.right);
            w32(76,node.child);
            w32(116,node.start);
            w64(120,node.size);
            return out;
        }
        std::vector<CfbNode> makeNodes(const std::map<std::string,Bytes>& streams) {
            std::vector<CfbNode> nodes;
            nodes.push_back({"Root Entry",5,-1,{}});
            std::unordered_map<std::string,std::uint32_t> storage {
                {
                    "",0
                }
            };
            for(const auto&[path,data]:streams) {
                std::size_t start=0;
                std::string prefix;
                std::uint32_t parent=0;
                while(true) {
                    auto slash=path.find('/',start);
                    if(slash==std::string::npos)break;
                    auto part=path.substr(start,slash-start);
                    prefix=prefix.empty()?part:prefix+"/"+part;
                    auto it=storage.find(prefix);
                    if(it==storage.end()) {
                        auto id=static_cast<std::uint32_t>(nodes.size());
                        nodes.push_back({part,1,static_cast<int>(parent),{}});
                        storage[prefix]=id;
                        parent=id;
                    } else parent=it->second;
                    start=slash+1;
                }
                nodes.push_back({path.substr(start),2,static_cast<int>(parent),data});
            }
            for(std::uint32_t i=0;i<nodes.size();++i)if(nodes[i].type==1||nodes[i].type==5)assignChildTree(nodes,i);
            return nodes;
        }
        Bytes buildCfb(const std::map<std::string,Bytes>& streams) {
            auto nodes=makeNodes(streams);
            std::vector<std::uint32_t>miniFat;
            Bytes miniStream;
            struct Reg {
                std::uint32_t node;
                Bytes data;
            };
            std::vector<Reg>regular;
            for(std::uint32_t i=1;i<nodes.size();++i) {
                auto&n=nodes[i];
                if(n.type!=2)continue;
                n.size=n.data.size();
                if(n.data.empty()) {
                    n.start=kEndOfChain;
                    continue;
                }
                if(n.data.size()<kMiniCutoff) {
                    auto first=static_cast<std::uint32_t>(miniFat.size());
                    auto count=(n.data.size()+63)/64;
                    n.start=first;
                    for(std::size_t s=0;s<count;++s) {
                        miniFat.push_back(s+1==count?kEndOfChain:first+static_cast<std::uint32_t>(s+1));
                        auto from=s*64,amount=std::min<std::size_t>(64,n.data.size()-from);
                        miniStream.insert(miniStream.end(),n.data.begin()+static_cast<std::ptrdiff_t>(from),n.data.begin()+static_cast<std::ptrdiff_t>(from+amount));
                        miniStream.resize(miniStream.size()+64-amount,0);
                    }
                } else regular.push_back({i,n.data});
            }
            nodes[0].size=miniStream.size();
            const auto dirBytes=((nodes.size()*128+511)/512)*512,dirSectors=dirBytes/512,rootMiniSectors=(miniStream.size()+511)/512,miniFatSectors=miniFat.empty()?0:(miniFat.size()*4+511)/512;
            std::size_t regularSectors=0;
            for(auto&b:regular)regularSectors+=(b.data.size()+511)/512;
            const auto nonFat=regularSectors+rootMiniSectors+dirSectors+miniFatSectors;
            std::size_t fatSectors=1,difatSectors=0;
            for(;;) {
                auto newDif=fatSectors>109?(fatSectors-109+126)/127:0;
                auto total=nonFat+fatSectors+newDif;
                auto newFat=(total+127)/128;
                if(newFat==fatSectors&&newDif==difatSectors)break;
                fatSectors=newFat;
                difatSectors=newDif;
            }
            const auto total=nonFat+fatSectors+difatSectors;
            std::vector<std::uint32_t>fat(fatSectors*128,kFreeSect);
            std::vector<std::array<unsigned char,512>>sectors(total);
            std::size_t next=0;
            auto place=[&](const Bytes&data,std::size_t count) {
                if(!count)return kEndOfChain;
                auto first=static_cast<std::uint32_t>(next);
                for(std::size_t s=0;s<count;++s) {
                    auto from=s*512,amount=std::min<std::size_t>(512,data.size()>from?data.size()-from:0);
                    if(amount)std::copy_n(data.begin()+static_cast<std::ptrdiff_t>(from),amount,sectors[next].begin());
                    fat[next]=s+1==count?kEndOfChain:static_cast<std::uint32_t>(next+1);
                    ++next;
                }
                return first;
            };
            for(auto&b:regular) {
                auto count=(b.data.size()+511)/512;
                nodes[b.node].start=place(b.data,count);
            }
            nodes[0].start=place(miniStream,rootMiniSectors);
            Bytes dir(dirBytes,0);
            for(std::size_t i=0;i<nodes.size();++i) {
                auto e=directoryEntry(nodes[i]);
                std::copy(e.begin(),e.end(),dir.begin()+static_cast<std::ptrdiff_t>(i*128));
            }
            auto firstDir=place(dir,dirSectors);
            Bytes mini(miniFatSectors*512,0xff);
            for(std::size_t i=0;i<miniFat.size();++i)for(unsigned s=0;s<32;s+=8)mini[i*4+s/8]=static_cast<unsigned char>(miniFat[i]>>s);
            auto firstMini=place(mini,miniFatSectors);
            const auto firstFat=next;
            for(std::size_t i=0;i<fatSectors;++i)fat[next+i]=kFatSect;
            next+=fatSectors;
            const auto firstDif=difatSectors?static_cast<std::uint32_t>(next):kEndOfChain;
            for(std::size_t i=0;i<difatSectors;++i)fat[next+i]=kDifSect;
            for(std::size_t f=0;f<fatSectors;++f) {
                auto&target=sectors[firstFat+f];
                for(std::size_t j=0;j<128;++j) {
                    auto value=fat[f*128+j];
                    for(unsigned s=0;s<32;s+=8)target[j*4+s/8]=static_cast<unsigned char>(value>>s);
                }
            }
            if(difatSectors) {
                std::size_t fatId=109;
                for(std::size_t d=0;d<difatSectors;++d) {
                    auto&target=sectors[firstDif+d];
                    target.fill(0xff);
                    for(std::size_t j=0;j<127&&fatId<fatSectors;++j,++fatId) {
                        auto value=static_cast<std::uint32_t>(firstFat+fatId);
                        for(unsigned s=0;s<32;s+=8)target[j*4+s/8]=static_cast<unsigned char>(value>>s);
                    }
                    auto nextDif=d+1<difatSectors?static_cast<std::uint32_t>(firstDif+d+1):kEndOfChain;
                    for(unsigned s=0;s<32;s+=8)target[508+s/8]=static_cast<unsigned char>(nextDif>>s);
                }
            }
            Bytes out(512,0);
            const unsigned char sig[8]= {
                0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1
            };
            std::copy(std::begin(sig),std::end(sig),out.begin());
            overwriteU16(out,24,0x003e);
            overwriteU16(out,26,3);
            overwriteU16(out,28,0xfffe);
            overwriteU16(out,30,9);
            overwriteU16(out,32,6);
            overwriteU32(out,40,0);
            overwriteU32(out,44,static_cast<std::uint32_t>(fatSectors));
            overwriteU32(out,48,firstDir);
            overwriteU32(out,56,4096);
            overwriteU32(out,60,firstMini);
            overwriteU32(out,64,static_cast<std::uint32_t>(miniFatSectors));
            overwriteU32(out,68,firstDif);
            overwriteU32(out,72,static_cast<std::uint32_t>(difatSectors));
            for(std::size_t i=0;i<109;++i)overwriteU32(out,76+i*4,i<fatSectors?static_cast<std::uint32_t>(firstFat+i):kFreeSect);
            for(auto&s:sectors)out.insert(out.end(),s.begin(),s.end());
            return out;
        }
        struct DirEntry {
            std::string name;
            std::uint8_t type {
            };
            std::uint32_t left {
                kFreeSect
            },
            right {
                kFreeSect
            },
            child {
                kFreeSect
            },
            start {
                kEndOfChain
            };
            std::uint64_t size {
            };
        };
        class CfbReader {
            public:explicit CfbReader(const Bytes&bytes):bytes_(bytes) {
                parse();
            }
            bool has(const std::string&path)const {
                return paths_.contains(lowerAscii(path));
            }
            Bytes stream(const std::string&path)const {
                auto it=paths_.find(lowerAscii(path));
                if(it==paths_.end())throw std::runtime_error("CFB stream not found: "+path);
                const auto&e=dir_.at(it->second);
                if(e.type!=2)throw std::runtime_error("CFB path is not a stream: "+path);
                if(!e.size)return {
                };
                return e.size<miniCutoff_?readMini(e.start,e.size):readRegular(e.start,e.size);
            }
            private:const Bytes&bytes_;
            std::size_t sectorSize_ {
                512
            },
            miniSectorSize_ {
                64
            },
            miniCutoff_ {
                4096
            };
            std::vector<std::uint32_t>fat_,miniFat_;
            std::vector<DirEntry>dir_;
            Bytes rootMini_;
            std::unordered_map<std::string,std::size_t>paths_;
            const unsigned char*sector(std::uint32_t id)const {
                auto off=512+static_cast<std::size_t>(id)*sectorSize_;
                if(off+sectorSize_>bytes_.size())throw std::runtime_error("CFB sector outside file");
                return bytes_.data()+off;
            }
            Bytes readRegular(std::uint32_t start,std::uint64_t expected)const {
                Bytes out;
                if(start==kEndOfChain)return out;
                auto id=start;
                std::size_t guard=0;
                while(id!=kEndOfChain&&id!=kFreeSect) {
                    if(id>=fat_.size())throw std::runtime_error("Invalid CFB FAT chain");
                    auto*p=sector(id);
                    out.insert(out.end(),p,p+sectorSize_);
                    id=fat_[id];
                    if(++guard>fat_.size())throw std::runtime_error("Cyclic CFB FAT chain");
                }
                if(expected<out.size())out.resize(static_cast<std::size_t>(expected));
                return out;
            }
            Bytes readMini(std::uint32_t start,std::uint64_t expected)const {
                Bytes out;
                auto id=start;
                std::size_t guard=0;
                while(id!=kEndOfChain&&id!=kFreeSect&&out.size()<expected) {
                    if(id>=miniFat_.size())throw std::runtime_error("Invalid CFB miniFAT chain");
                    auto off=static_cast<std::size_t>(id)*miniSectorSize_;
                    if(off+miniSectorSize_>rootMini_.size())throw std::runtime_error("CFB mini sector outside root stream");
                    out.insert(out.end(),rootMini_.begin()+static_cast<std::ptrdiff_t>(off),rootMini_.begin()+static_cast<std::ptrdiff_t>(off+miniSectorSize_));
                    id=miniFat_[id];
                    if(++guard>miniFat_.size())throw std::runtime_error("Cyclic CFB miniFAT chain");
                }
                if(expected<out.size())out.resize(static_cast<std::size_t>(expected));
                return out;
            }
            void walk(std::uint32_t id,const std::string&prefix,std::vector<bool>&visited) {
                if(id==kFreeSect||id>=dir_.size()||visited[id])return;
                visited[id]=true;
                auto e=dir_[id];
                walk(e.left,prefix,visited);
                auto path=prefix.empty()?e.name:prefix+"/"+e.name;
                paths_[lowerAscii(path)]=id;
                if(e.type==1||e.type==5)walk(e.child,path=="Root Entry"?"":path,visited);
                walk(e.right,prefix,visited);
            }
            void parse() {
                if(bytes_.size()<512||!hasCompoundFileSignature(bytes_.data(),bytes_.size()))throw std::runtime_error("Invalid CFB signature");
                const auto sectorShift=getU16(bytes_.data()+30),miniShift=getU16(bytes_.data()+32);
                if(sectorShift!=9||miniShift!=6)throw std::runtime_error("Unsupported CFB sector geometry");
                sectorSize_=512;
                miniSectorSize_=64;
                miniCutoff_=getU32(bytes_.data()+56);
                if(miniCutoff_==0||miniCutoff_>16*1024*1024)throw std::runtime_error("Invalid CFB mini-stream cutoff");
                const auto maxSectors=(bytes_.size()-512)/sectorSize_;
                auto fatCount=getU32(bytes_.data()+44),firstDir=getU32(bytes_.data()+48),firstMini=getU32(bytes_.data()+60),miniCount=getU32(bytes_.data()+64),firstDif=getU32(bytes_.data()+68),difCount=getU32(bytes_.data()+72);
                if(fatCount>maxSectors||difCount>maxSectors||miniCount>maxSectors)throw std::runtime_error("CFB allocation table counts exceed file size");
                std::vector<std::uint32_t>fatSectors;
                for(std::size_t i=0;i<109&&fatSectors.size()<fatCount;++i) {
                    auto id=getU32(bytes_.data()+76+i*4);
                    if(id!=kFreeSect)fatSectors.push_back(id);
                }
                auto dif=firstDif;
                std::unordered_set<std::uint32_t>seenDif;
                for(std::uint32_t d=0;d<difCount&&dif!=kEndOfChain;++d) {
                    if(!seenDif.insert(dif).second)throw std::runtime_error("Cyclic CFB DIFAT chain");
                    auto*p=sector(dif);
                    for(std::size_t i=0;i<127&&fatSectors.size()<fatCount;++i) {
                        auto id=getU32(p+i*4);
                        if(id!=kFreeSect)fatSectors.push_back(id);
                    }
                    dif=getU32(p+508);
                }
                if(fatSectors.size()<fatCount)throw std::runtime_error("Incomplete CFB FAT list");
                for(auto id:fatSectors) {
                    auto*p=sector(id);
                    for(std::size_t i=0;i<128;++i)fat_.push_back(getU32(p+i*4));
                }
                auto dirBytes=readRegular(firstDir,std::numeric_limits<std::uint64_t>::max());
                for(std::size_t off=0;off+128<=dirBytes.size();off+=128) {
                    auto*p=dirBytes.data()+off;
                    auto nameBytes=getU16(p+64);
                    DirEntry e;
                    if(nameBytes>=2&&nameBytes<=64)for(std::size_t i=0;i+1<static_cast<std::size_t>(nameBytes-2);i+=2)e.name.push_back(static_cast<char>(p[i]));
                    e.type=p[66];
                    e.left=getU32(p+68);
                    e.right=getU32(p+72);
                    e.child=getU32(p+76);
                    e.start=getU32(p+116);
                    e.size=getU64(p+120);
                    dir_.push_back(std::move(e));
                }
                if(dir_.empty()||dir_[0].type!=5)throw std::runtime_error("CFB root missing");
                rootMini_=readRegular(dir_[0].start,dir_[0].size);
                if(miniCount&&firstMini!=kEndOfChain) {
                    auto bytes=readRegular(firstMini,static_cast<std::uint64_t>(miniCount)*512);
                    for(std::size_t i=0;i+4<=bytes.size();i+=4)miniFat_.push_back(getU32(bytes.data()+i));
                }
                std::vector<bool>visited(dir_.size(),false);
                paths_["root entry"]=0;
                walk(dir_[0].child,"",visited);
            }
        };
        // ---------- DataSpaces static structures, generated from MS-OFFCRYPTO fields ----------
        void putUnicodeLpP4(Bytes& out,std::string_view text) {
            const auto bytes=text.size()*2;
            putU32(out,static_cast<std::uint32_t>(bytes));
            for(char raw:text) {
                const auto c = static_cast<unsigned char>(raw);
                out.push_back(c);
                out.push_back(0);
            }
            if((4+bytes)%4!=0) {
                out.push_back(0);
                out.push_back(0);
            }
        }
        Bytes dataSpacesVersion() {
            Bytes out;
            putUnicodeLpP4(out,"Microsoft.Container.DataSpaces");
            putU32(out,1);
            putU32(out,1);
            putU32(out,1);
            return out;
        }
        Bytes dataSpaceMap() {
            Bytes out;
            putU32(out,8);
            putU32(out,1);
            Bytes entry;
            putU32(entry,1);
            putU32(entry,0);
            putUnicodeLpP4(entry,"EncryptedPackage");
            putUnicodeLpP4(entry,"StrongEncryptionDataSpace");
            putU32(out,static_cast<std::uint32_t>(entry.size()+4));
            out.insert(out.end(),entry.begin(),entry.end());
            return out;
        }
        Bytes strongEncryptionDataSpace() {
            Bytes out;
            putU32(out,8);
            putU32(out,1);
            putUnicodeLpP4(out,"StrongEncryptionTransform");
            return out;
        }
        Bytes strongEncryptionPrimary() {
            Bytes out;
            putU32(out,0x58);
            putU32(out,1);
            putUnicodeLpP4(out,"{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}");
            putUnicodeLpP4(out,"Microsoft.Container.EncryptionTransform");
            putU32(out,1);
            putU32(out,1);
            putU32(out,1);
            putU32(out,0);
            putU32(out,0);
            putU32(out,0);
            putU32(out,4);
            return out;
        }
        std::string attr(std::string_view tag,std::string_view name) {
            auto pos=tag.find(name);
            while(pos!=std::string_view::npos) {
                const auto before=pos==0?' ':tag[pos-1];
                if(std::isspace(static_cast<unsigned char>(before))||before=='<') {
                    auto p=pos+name.size();
                    while(p<tag.size()&&std::isspace(static_cast<unsigned char>(tag[p])))++p;
                    if(p<tag.size()&&tag[p]=='=') {
                        ++p;
                        while(p<tag.size()&&std::isspace(static_cast<unsigned char>(tag[p])))++p;
                        if(p<tag.size()&&(tag[p]=='\''||tag[p]=='"')) {
                            char q=tag[p++];
                            auto end=tag.find(q,p);
                            if(end!=std::string_view::npos)return std::string(tag.substr(p,end-p));
                        }
                    }
                }
                pos=tag.find(name,pos+1);
            }
            return {
            };
        }
        std::string_view startTag(std::string_view xml,std::string_view needle) {
            auto pos=xml.find(needle);
            if(pos==std::string_view::npos)return {
            };
            auto end=xml.find('>',pos);
            if(end==std::string_view::npos)return {
            };
            return xml.substr(pos,end-pos+1);
        }
        struct StandardInfo {
            std::uint16_t major {
            },
            minor {
            };
            std::uint32_t flags {
            },
            algId {
            },
            algIdHash {
            },
            keyBits {
            },
            providerType {
            };
            Bytes salt, encryptedVerifier, encryptedVerifierHash;
            std::uint32_t verifierHashSize {
            };
        };
        StandardInfo parseStandard(const Bytes& encryptionInfo, bool requireSupported=true) {
            if (encryptionInfo.size() < 12) throw std::runtime_error("Standard EncryptionInfo is truncated");
            StandardInfo i;
            i.major=getU16(encryptionInfo.data());
            i.minor=getU16(encryptionInfo.data()+2);
            if (!((i.major==2||i.major==3||i.major==4) && i.minor==2)) throw std::runtime_error("Unsupported Standard EncryptionInfo version");
            i.flags=getU32(encryptionInfo.data()+4);
            const auto headerSize=getU32(encryptionInfo.data()+8);
            if (headerSize < 32 || static_cast<std::uint64_t>(12)+headerSize+4 > encryptionInfo.size()) throw std::runtime_error("Standard EncryptionHeader is truncated");
            const auto* h=encryptionInfo.data()+12;
            const auto headerFlags=getU32(h);
            (void)headerFlags;
            i.algId=getU32(h+8);
            i.algIdHash=getU32(h+12);
            i.keyBits=getU32(h+16);
            i.providerType=getU32(h+20);
            const bool aesAlg=i.algId==0x660eu||i.algId==0x660fu||i.algId==0x6610u;
            const bool keyOk=(i.keyBits==128||i.keyBits==192||i.keyBits==256) && ((i.algId==0x660e&&i.keyBits==128)||(i.algId==0x660f&&i.keyBits==192)||(i.algId==0x6610&&i.keyBits==256));
            const bool supported=aesAlg&&keyOk&&i.algIdHash==0x8004u&&i.providerType==0x18u;
            if (requireSupported && !supported) throw std::runtime_error("Unsupported Standard encryption algorithm (XL++ supports AES-128/192/256 with SHA-1)");
            std::size_t off=12+headerSize;
            if(off+4>encryptionInfo.size())throw std::runtime_error("Standard EncryptionVerifier is truncated");
            const auto saltSize=getU32(encryptionInfo.data()+off);
            off+=4;
            if(saltSize!=16 || off+saltSize+16+4>encryptionInfo.size())throw std::runtime_error("Unsupported Standard encryption salt/verifier size");
            i.salt.assign(encryptionInfo.begin()+static_cast<std::ptrdiff_t>(off),encryptionInfo.begin()+static_cast<std::ptrdiff_t>(off+saltSize));
            off+=saltSize;
            i.encryptedVerifier.assign(encryptionInfo.begin()+static_cast<std::ptrdiff_t>(off),encryptionInfo.begin()+static_cast<std::ptrdiff_t>(off+16));
            off+=16;
            i.verifierHashSize=getU32(encryptionInfo.data()+off);
            off+=4;
            if(i.verifierHashSize==0||i.verifierHashSize>64)throw std::runtime_error("Invalid Standard verifier hash size");
            const auto encryptedHashSize=((static_cast<std::size_t>(i.verifierHashSize)+15)/16)*16;
            if(off+encryptedHashSize>encryptionInfo.size())throw std::runtime_error("Standard encrypted verifier hash is truncated");
            i.encryptedVerifierHash.assign(encryptionInfo.begin()+static_cast<std::ptrdiff_t>(off),encryptionInfo.begin()+static_cast<std::ptrdiff_t>(off+encryptedHashSize));
            return i;
        }
        Bytes deriveStandardKey(const StandardInfo& info,const std::string& password) {
            Bytes initial=info.salt;
            auto passwordBytes=utf16LePassword(password);
            initial.insert(initial.end(),passwordBytes.begin(),passwordBytes.end());
            auto h=sha1(initial);
            for(std::uint32_t iteration=0;iteration<50000;++iteration) {
                auto input=little32(iteration);
                input.insert(input.end(),h.begin(),h.end());
                h=sha1(input);
            }
            auto finalInput=h;
            auto block=little32(0);
            finalInput.insert(finalInput.end(),block.begin(),block.end());
            auto hfinal=sha1(finalInput);
            Bytes pad36(64,0x36),pad5c(64,0x5c);
            for(std::size_t n=0;n<hfinal.size();++n) {
                pad36[n]^=hfinal[n];
                pad5c[n]^=hfinal[n];
            }
            auto x1=sha1(pad36),x2=sha1(pad5c);
            x1.insert(x1.end(),x2.begin(),x2.end());
            const auto keyBytes=static_cast<std::size_t>(info.keyBits/8);
            if(x1.size()<keyBytes)throw std::runtime_error("Standard derived key is too short");
            x1.resize(keyBytes);
            return x1;
        }
        Bytes recoverStandardKey(const StandardInfo& info,const std::string& password) {
            auto key=deriveStandardKey(info,password);
            auto verifier=aesEcbDecrypt(info.encryptedVerifier,key);
            auto expected=sha1(verifier);
            auto decryptedHash=aesEcbDecrypt(info.encryptedVerifierHash,key);
            if (decryptedHash.size() < expected.size()) throw std::runtime_error("Standard verifier hash is truncated");
            decryptedHash.resize(expected.size());
            if (!constantTimeEqual(expected, decryptedHash)) throw std::runtime_error("Incorrect workbook password");
            return key;
        }
        Bytes decryptStandardPackageStream(const Bytes& stream,const Bytes& key) {
            if (stream.size() < 8) throw std::runtime_error("EncryptedPackage stream is truncated");
            const auto size = getU64(stream.data());
            if(size>std::numeric_limits<std::size_t>::max())throw std::runtime_error("EncryptedPackage is too large");
            Bytes cipher(stream.begin()+8,stream.end());
            if(size>cipher.size())throw std::runtime_error("Standard EncryptedPackage plaintext size is invalid");
            if(cipher.size()%16)throw std::runtime_error("Standard EncryptedPackage is not AES block aligned");
            auto plain=aesEcbDecrypt(cipher,key);
            if (size > plain.size()) throw std::runtime_error("Standard EncryptedPackage plaintext size exceeds decrypted data");
            plain.resize(static_cast<std::size_t>(size));
            return plain;
        }
        struct AgileInfo {
            Bytes keySalt,passwordSalt,encryptedVerifierInput,encryptedVerifierHash,encryptedKey,encryptedHmacKey,encryptedHmacValue;
            std::uint32_t spinCount {
            };
        };
        AgileInfo parseAgile(const Bytes& encryptionInfo,bool requireSupported=true) {
            if(encryptionInfo.size()<9||getU16(encryptionInfo.data())!=4||getU16(encryptionInfo.data()+2)!=4||getU32(encryptionInfo.data()+4)!=0x40)throw std::runtime_error("Unsupported EncryptionInfo version");
            std::string xml(encryptionInfo.begin()+8,encryptionInfo.end());
            auto kd=startTag(xml,"<keyData");
            auto ek=startTag(xml,"<p:encryptedKey");
            if(ek.empty())ek=startTag(xml,"<encryptedKey");
            auto di=startTag(xml,"<dataIntegrity");
            if(kd.empty()||ek.empty())throw std::runtime_error("Agile EncryptionInfo is incomplete");
            const auto cipher=attr(kd,"cipherAlgorithm"),chain=attr(kd,"cipherChaining"),hash=attr(kd,"hashAlgorithm"),bits=attr(kd,"keyBits");
            const bool supported=cipher=="AES"&&chain=="ChainingModeCBC"&&hash=="SHA512"&&bits=="256";
            if(requireSupported&&!supported)throw std::runtime_error("Unsupported Agile encryption algorithm (XL++ supports AES-256-CBC/SHA-512)");
            AgileInfo i;
            i.keySalt=base64Decode(attr(kd,"saltValue"));
            i.passwordSalt=base64Decode(attr(ek,"saltValue"));
            i.encryptedVerifierInput=base64Decode(attr(ek,"encryptedVerifierHashInput"));
            i.encryptedVerifierHash=base64Decode(attr(ek,"encryptedVerifierHashValue"));
            i.encryptedKey=base64Decode(attr(ek,"encryptedKeyValue"));
            if(!di.empty()) {
                i.encryptedHmacKey=base64Decode(attr(di,"encryptedHmacKey"));
                i.encryptedHmacValue=base64Decode(attr(di,"encryptedHmacValue"));
            }
            const auto spin=attr(ek,"spinCount");
            i.spinCount=spin.empty()?0:static_cast<std::uint32_t>(std::stoul(spin));
            if(i.spinCount>10'000'000)throw std::runtime_error("Agile encryption spin count exceeds safety limit");
            if(i.keySalt.empty()||i.passwordSalt.empty())throw std::runtime_error("Agile encryption salt is missing");
            return i;
        }
        Bytes recoverSecret(const AgileInfo& info,const std::string&password) {
            auto h=deriveSpinHash(info.passwordSalt,password,info.spinCount);
            auto iv=deriveIv(info.passwordSalt,std::nullopt);
            auto verifier=aesCbcDecrypt(info.encryptedVerifierInput,deriveKey(h,asBytes(kVerifierInputBlock)),iv);
            if(verifier.size()<16)throw std::runtime_error("Invalid encrypted verifier");
            verifier.resize(16);
            auto verifierHash=aesCbcDecrypt(info.encryptedVerifierHash,deriveKey(h,asBytes(kVerifierHashBlock)),iv);
            auto expected=sha512(verifier);
            verifierHash.resize(expected.size());
            if(!constantTimeEqual(expected,verifierHash))throw std::runtime_error("Incorrect workbook password");
            auto secret=aesCbcDecrypt(info.encryptedKey,deriveKey(h,asBytes(kEncryptedKeyBlock)),iv);
            if(secret.size()<32)throw std::runtime_error("Invalid encrypted package key");
            secret.resize(32);
            return secret;
        }
        Bytes encryptPackageStream(const Bytes& package,const Bytes& secret,const Bytes& salt) {
            Bytes stream;
            putU64(stream,package.size());
            std::size_t offset=0;
            std::uint32_t segment=0;
            while(offset<package.size()) {
                auto count=std::min<std::size_t>(4096,package.size()-offset);
                Bytes plain(package.begin()+static_cast<std::ptrdiff_t>(offset),package.begin()+static_cast<std::ptrdiff_t>(offset+count));
                plain=zeroPad(std::move(plain));
                auto iv=deriveIv(salt,little32(segment++));
                auto cipher=aesCbcEncrypt(plain,secret,iv);
                stream.insert(stream.end(),cipher.begin(),cipher.end());
                offset+=count;
            }
            return stream;
        }
        Bytes decryptPackageStream(const Bytes& stream,const Bytes& secret,const Bytes& salt) {
            if(stream.size()<8)throw std::runtime_error("EncryptedPackage stream is truncated");
            auto size=getU64(stream.data());
            if(size>std::numeric_limits<std::size_t>::max()||size>stream.size()-8)throw std::runtime_error("EncryptedPackage plaintext size is invalid");
            Bytes out;
            out.reserve(static_cast<std::size_t>(size));
            std::size_t cipherOffset=8,remaining=static_cast<std::size_t>(size);
            std::uint32_t segment=0;
            while(remaining) {
                auto plainCount=std::min<std::size_t>(4096,remaining);
                auto cipherCount=((plainCount+15)/16)*16;
                if(cipherOffset+cipherCount>stream.size())throw std::runtime_error("EncryptedPackage segment is truncated");
                Bytes cipher(stream.begin()+static_cast<std::ptrdiff_t>(cipherOffset),stream.begin()+static_cast<std::ptrdiff_t>(cipherOffset+cipherCount));
                auto plain=aesCbcDecrypt(cipher,secret,deriveIv(salt,little32(segment++)));
                out.insert(out.end(),plain.begin(),plain.begin()+static_cast<std::ptrdiff_t>(plainCount));
                cipherOffset+=cipherCount;
                remaining-=plainCount;
            }
            return out;
        }
        std::string agileXml(const Bytes& keySalt,const Bytes& passwordSalt,std::uint32_t spin,const Bytes& verifierInput,const Bytes& verifierHash,const Bytes& encryptedKey,const Bytes& hmacKey,const Bytes& hmacValue) {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" "<encryption xmlns=\"http://schemas.microsoft.com/office/2006/encryption\" xmlns:p=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">" "<keyData saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\""+base64Encode(keySalt)+"\"/>" "<dataIntegrity encryptedHmacKey=\""+base64Encode(hmacKey)+"\" encryptedHmacValue=\""+base64Encode(hmacValue)+"\"/>" "<keyEncryptors><keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\"><p:encryptedKey spinCount=\""+std::to_string(spin)+"\" saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\""+base64Encode(passwordSalt)+"\" encryptedVerifierHashInput=\""+base64Encode(verifierInput)+"\" encryptedVerifierHashValue=\""+base64Encode(verifierHash)+"\" encryptedKeyValue=\""+base64Encode(encryptedKey)+"\"/></keyEncryptor></keyEncryptors></encryption>";
        }
    }
    // namespace
    bool hasCompoundFileSignature(const unsigned char* data,std::size_t size) noexcept {
        static const unsigned char sig[8]= {
            0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1
        };
        return size>=8&&std::equal(std::begin(sig),std::end(sig),data);
    }
    std::vector<unsigned char> encryptStandardOfficePackage(const std::vector<unsigned char>& package,const std::string& passwordUtf8,std::uint32_t keyBits) {
        if(passwordUtf8.empty())throw std::invalid_argument("Encryption password cannot be empty");
        std::uint32_t algId=0;
        if(keyBits==128)algId=0x660e;
        else if(keyBits==192)algId=0x660f;
        else if(keyBits==256)algId=0x6610;
        else throw std::invalid_argument("Standard AES encryption key size must be 128, 192 or 256 bits");
        StandardInfo info;
        info.major=3;
        info.minor=2;
        info.flags=0x24;
        info.algId=algId;
        info.algIdHash=0x8004;
        info.keyBits=keyBits;
        info.providerType=0x18;
        info.salt=randomBytes(16);
        info.verifierHashSize=20;
        const auto key=deriveStandardKey(info,passwordUtf8);
        const auto verifier=randomBytes(16);
        info.encryptedVerifier=aesEcbEncrypt(verifier,key);
        info.encryptedVerifierHash=aesEcbEncrypt(zeroPad(sha1(verifier)),key);
        Bytes header;
        putU32(header,0x24);
        putU32(header,0);
        putU32(header,algId);
        putU32(header,0x8004);
        putU32(header,keyBits);
        putU32(header,0x18);
        putU32(header,0);
        putU32(header,0);
        const auto provider=utf16LeName("Microsoft Enhanced RSA and AES Cryptographic Provider");
        header.insert(header.end(),provider.begin(),provider.end());
        Bytes encryptionInfo;
        putU16(encryptionInfo,3);
        putU16(encryptionInfo,2);
        putU32(encryptionInfo,0x24);
        putU32(encryptionInfo,static_cast<std::uint32_t>(header.size()));
        encryptionInfo.insert(encryptionInfo.end(),header.begin(),header.end());
        putU32(encryptionInfo,16);
        encryptionInfo.insert(encryptionInfo.end(),info.salt.begin(),info.salt.end());
        encryptionInfo.insert(encryptionInfo.end(),info.encryptedVerifier.begin(),info.encryptedVerifier.end());
        putU32(encryptionInfo,20);
        encryptionInfo.insert(encryptionInfo.end(),info.encryptedVerifierHash.begin(),info.encryptedVerifierHash.end());
        Bytes encryptedPackage;
        putU64(encryptedPackage,package.size());
        auto plain=zeroPad(Bytes(package.begin(),package.end()));
        auto cipher=aesEcbEncrypt(plain,key);
        encryptedPackage.insert(encryptedPackage.end(),cipher.begin(),cipher.end());
        std::map<std::string,Bytes> streams;
        streams["EncryptedPackage"]=std::move(encryptedPackage);
        streams[std::string(1,'\x06')+"DataSpaces/Version"]=dataSpacesVersion();
        streams[std::string(1,'\x06')+"DataSpaces/DataSpaceMap"]=dataSpaceMap();
        streams[std::string(1,'\x06')+"DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace"]=strongEncryptionDataSpace();
        streams[std::string(1,'\x06')+"DataSpaces/TransformInfo/StrongEncryptionTransform/"+std::string(1,'\x06')+"Primary"]=strongEncryptionPrimary();
        streams["EncryptionInfo"]=std::move(encryptionInfo);
        return buildCfb(streams);
    }
    std::vector<unsigned char> encryptAgileOfficePackage(const std::vector<unsigned char>& package,const std::string& passwordUtf8,std::uint32_t spinCount) {
        if (passwordUtf8.empty()) throw std::invalid_argument("Encryption password cannot be empty");
        if (spinCount > 10'000'000) throw std::invalid_argument("Agile encryption spin count exceeds 10,000,000");
        const auto keySalt=randomBytes(16),passwordSalt=randomBytes(16),verifier=randomBytes(16),secret=randomBytes(32);
        auto spin=deriveSpinHash(passwordSalt,passwordUtf8,spinCount),passwordIv=deriveIv(passwordSalt,std::nullopt);
        auto encryptedVerifier=aesCbcEncrypt(verifier,deriveKey(spin,asBytes(kVerifierInputBlock)),passwordIv);
        auto encryptedVerifierHash=aesCbcEncrypt(zeroPad(sha512(verifier)),deriveKey(spin,asBytes(kVerifierHashBlock)),passwordIv);
        auto encryptedKey=aesCbcEncrypt(secret,deriveKey(spin,asBytes(kEncryptedKeyBlock)),passwordIv);
        auto encryptedPackage=encryptPackageStream(package,secret,keySalt);
        auto hmacSecret=randomBytes(64);
        auto hmac=hmacSha512(hmacSecret,encryptedPackage);
        auto encryptedHmacKey=aesCbcEncrypt(hmacSecret,secret,deriveIv(keySalt,asBytes(kHmacKeyBlock)));
        auto encryptedHmacValue=aesCbcEncrypt(zeroPad(hmac),secret,deriveIv(keySalt,asBytes(kHmacValueBlock)));
        auto xml=agileXml(keySalt,passwordSalt,spinCount,encryptedVerifier,encryptedVerifierHash,encryptedKey,encryptedHmacKey,encryptedHmacValue);
        Bytes encryptionInfo;
        putU16(encryptionInfo,4);
        putU16(encryptionInfo,4);
        putU32(encryptionInfo,0x40);
        encryptionInfo.insert(encryptionInfo.end(),xml.begin(),xml.end());
        std::map<std::string,Bytes> streams;
        streams["EncryptedPackage"]=std::move(encryptedPackage);
        streams[std::string(1,'\x06')+"DataSpaces/Version"]=dataSpacesVersion();
        streams[std::string(1,'\x06')+"DataSpaces/DataSpaceMap"]=dataSpaceMap();
        streams[std::string(1,'\x06')+"DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace"]=strongEncryptionDataSpace();
        streams[std::string(1,'\x06')+"DataSpaces/TransformInfo/StrongEncryptionTransform/"+std::string(1,'\x06')+"Primary"]=strongEncryptionPrimary();
        streams["EncryptionInfo"]=std::move(encryptionInfo);
        return buildCfb(streams);
    }
    std::vector<unsigned char> decryptAgileOfficePackage(const std::vector<unsigned char>& compoundFile,const std::string& passwordUtf8,bool verifyIntegrity) {
        CfbReader cfb(compoundFile);
        if(!cfb.has("EncryptionInfo")||!cfb.has("EncryptedPackage"))throw std::runtime_error("CFB file is not an encrypted ECMA-376 package");
        auto info=parseAgile(cfb.stream("EncryptionInfo"));
        auto secret=recoverSecret(info,passwordUtf8);
        auto encryptedPackage=cfb.stream("EncryptedPackage");
        if(verifyIntegrity&&!info.encryptedHmacKey.empty()&&!info.encryptedHmacValue.empty()) {
            auto hmacKey=aesCbcDecrypt(info.encryptedHmacKey,secret,deriveIv(info.keySalt,asBytes(kHmacKeyBlock)));
            if(hmacKey.size()<64)throw std::runtime_error("Encrypted HMAC key is truncated");
            hmacKey.resize(64);
            auto expected=aesCbcDecrypt(info.encryptedHmacValue,secret,deriveIv(info.keySalt,asBytes(kHmacValueBlock)));
            auto actual=hmacSha512(hmacKey,encryptedPackage);
            expected.resize(actual.size());
            if(!constantTimeEqual(actual,expected))throw std::runtime_error("Encrypted workbook integrity check failed");
        }
        return decryptPackageStream(encryptedPackage,secret,info.keySalt);
    }
    std::vector<unsigned char> decryptOfficePackage(const std::vector<unsigned char>& compoundFile,const std::string& passwordUtf8,bool verifyIntegrity) {
        CfbReader cfb(compoundFile);
        if(!cfb.has("EncryptionInfo")||!cfb.has("EncryptedPackage"))throw std::runtime_error("CFB file is not an encrypted ECMA-376 package");
        auto bytes=cfb.stream("EncryptionInfo");
        if(bytes.size()<4)throw std::runtime_error("EncryptionInfo is truncated");
        const auto major=getU16(bytes.data()),minor=getU16(bytes.data()+2);
        if(major==4&&minor==4)return decryptAgileOfficePackage(compoundFile,passwordUtf8,verifyIntegrity);
        if((major==2||major==3||major==4)&&minor==2) {
            auto info=parseStandard(bytes);
            auto key=recoverStandardKey(info,passwordUtf8);
            return decryptStandardPackageStream(cfb.stream("EncryptedPackage"),key);
        }
        throw std::runtime_error("Unsupported Office encryption mode");
    }
    OfficeEncryptionInfo inspectOfficeEncryptionBytes(const std::vector<unsigned char>& compoundFile) {
        OfficeEncryptionInfo result;
        if(!hasCompoundFileSignature(compoundFile.data(),compoundFile.size()))return result;
        result.encrypted=true;
        CfbReader cfb(compoundFile);
        if(!cfb.has("EncryptionInfo")||!cfb.has("EncryptedPackage")) {
            result.mode=OfficeEncryptionMode::Unsupported;
            return result;
        }
        auto bytes=cfb.stream("EncryptionInfo");
        if(bytes.size()<4) {
            result.mode=OfficeEncryptionMode::Unsupported;
            return result;
        }
        const auto major=getU16(bytes.data()),minor=getU16(bytes.data()+2);
        if(major==4&&minor==4) {
            if(bytes.size()<8) {
                result.mode=OfficeEncryptionMode::Unsupported;
                return result;
            }
            std::string xml(bytes.begin()+8,bytes.end());
            auto kd=startTag(xml,"<keyData");
            auto ek=startTag(xml,"<p:encryptedKey");
            result.cipherAlgorithm=attr(kd,"cipherAlgorithm");
            result.hashAlgorithm=attr(kd,"hashAlgorithm");
            auto bits=attr(kd,"keyBits"),spin=attr(ek,"spinCount");
            if(!bits.empty())result.keyBits=static_cast<std::uint32_t>(std::stoul(bits));
            if(!spin.empty())result.spinCount=static_cast<std::uint32_t>(std::stoul(spin));
            result.supported=result.cipherAlgorithm=="AES"&&attr(kd,"cipherChaining")=="ChainingModeCBC"&&result.hashAlgorithm=="SHA512"&&result.keyBits==256;
            result.mode=result.supported?OfficeEncryptionMode::AgileAes256Sha512:OfficeEncryptionMode::Unsupported;
            return result;
        }
        if((major==2||major==3||major==4)&&minor==2) {
            try {
                auto info=parseStandard(bytes,false);
                result.keyBits=info.keyBits;
                result.spinCount=50000;
                result.cipherAlgorithm=(info.algId==0x660e||info.algId==0x660f||info.algId==0x6610)?"AES":"Unknown";
                result.hashAlgorithm=info.algIdHash==0x8004?"SHA1":"Unknown";
                result.supported=result.cipherAlgorithm=="AES"&&result.hashAlgorithm=="SHA1"&&(result.keyBits==128||result.keyBits==192||result.keyBits==256)&&info.providerType==0x18;
                result.mode=result.supported?OfficeEncryptionMode::StandardAesSha1:OfficeEncryptionMode::Unsupported;
                return result;
            } catch (...) {
                result.mode=OfficeEncryptionMode::Unsupported;
                return result;
            }
        }
        result.mode=OfficeEncryptionMode::Unsupported;
        return result;
    }
}
// namespace xlpp::internal
namespace xlpp {
    bool looksLikeEncryptedOfficeFile(const std::filesystem::path& path) {
        std::ifstream in(path,std::ios::binary);
        std::array<unsigned char,8>head {
        };
        in.read(reinterpret_cast<char*>(head.data()),static_cast<std::streamsize>(head.size()));
        return internal::hasCompoundFileSignature(head.data(),static_cast<std::size_t>(in.gcount()));
    }
    OfficeEncryptionInfo inspectOfficeEncryption(const std::filesystem::path& path) {
        std::ifstream in(path,std::ios::binary);
        if(!in)throw std::runtime_error("Cannot open file: "+path.string());
        std::vector<unsigned char>bytes(std::istreambuf_iterator<char>(in),{});
        return internal::inspectOfficeEncryptionBytes(bytes);
    }
}
// namespace xlpp
