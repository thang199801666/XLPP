#pragma once
#include <cmath>
#include <cstdio>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

namespace xlpp {

// Calendar date and wall-clock time. Seconds are fractional so sub-second
// precision survives round-trips through the Excel serial representation.
struct DateTime {
    int year = 1900;
    int month = 1;
    int day = 1;
    int hour = 0;
    int minute = 0;
    double second = 0.0;

    friend bool operator==(const DateTime&, const DateTime&) noexcept = default;
    friend bool operator!=(const DateTime&, const DateTime&) noexcept = default;
};

// Days since 1970-01-01 for a civil (proleptic Gregorian) date. Port of the
// well-known days_from_civil algorithm.
inline long long daysFromCivil(int year, int month, int day) noexcept {
    year -= month <= 2;
    const long long era = (year >= 0 ? year : year - 399) / 400;
    const unsigned yearOfEra = static_cast<unsigned>(year - era * 400);
    const int monthAdjusted = month + (month > 2 ? -3 : 9);
    const unsigned dayOfYear = static_cast<unsigned>((153 * monthAdjusted + 2) / 5 + day - 1);
    const unsigned dayOfEra = yearOfEra * 365 + yearOfEra / 4 - yearOfEra / 100 + dayOfYear;
    return era * 146097 + static_cast<long long>(dayOfEra) - 719468;
}

inline void civilFromDays(long long days, int& year, int& month, int& day) noexcept {
    days += 719468;
    const long long era = (days >= 0 ? days : days - 146096) / 146097;
    const unsigned dayOfEra = static_cast<unsigned>(days - era * 146097);
    const unsigned yearOfEra = (dayOfEra - dayOfEra / 1460 + dayOfEra / 36524 - dayOfEra / 146096) / 365;
    const long long yearOfCycle = static_cast<long long>(yearOfEra) + era * 400;
    const unsigned dayOfYear = dayOfEra - (365 * yearOfEra + yearOfEra / 4 - yearOfEra / 100);
    const unsigned monthPrime = (5 * dayOfYear + 2) / 153;
    day = static_cast<int>(dayOfYear - (153 * monthPrime + 2) / 5 + 1);
    month = static_cast<int>(monthPrime) + (monthPrime < 10 ? 3 : -9);
    year = static_cast<int>(yearOfCycle + (month <= 2));
}

// Converts a calendar date/time to the Excel serial day number. The 1900 system
// counts from 1899-12-31 with a phantom 1900-02-29 (serial 60), matching Excel's
// leap-year bug; the 1904 system counts from 1904-01-01 with no such bug.
inline double toExcelSerial(const DateTime& value, bool date1904 = false) noexcept {
    const long long days = daysFromCivil(value.year, value.month, value.day);
    const long long epoch = date1904 ? daysFromCivil(1904, 1, 1) : daysFromCivil(1899, 12, 31);
    long long serial = days - epoch;
    if (!date1904 && serial >= 60) ++serial;
    const double fraction = (value.hour * 3600.0 + value.minute * 60.0 + value.second) / 86400.0;
    return static_cast<double>(serial) + fraction;
}

inline DateTime fromExcelSerial(double serial, bool date1904 = false) noexcept {
    const long long whole = static_cast<long long>(std::floor(serial));
    const double fraction = serial - static_cast<double>(whole);
    long long days = whole;
    if (!date1904 && days >= 60) --days;
    const long long epoch = date1904 ? daysFromCivil(1904, 1, 1) : daysFromCivil(1899, 12, 31);
    DateTime result;
    civilFromDays(epoch + days, result.year, result.month, result.day);
    double totalSeconds = fraction * 86400.0;
    totalSeconds = std::round(totalSeconds * 1000.0) / 1000.0;
    const long long wholeSeconds = static_cast<long long>(totalSeconds);
    const double subsecond = totalSeconds - static_cast<double>(wholeSeconds);
    result.hour = static_cast<int>((wholeSeconds / 3600) % 24);
    result.minute = static_cast<int>((wholeSeconds / 60) % 60);
    result.second = static_cast<double>(wholeSeconds % 60) + subsecond;
    return result;
}

// True when an Excel number-format code renders values as dates or times, used
// to decide whether a numeric cell should be exposed as a DateTime. Checks for
// built-in date/time format IDs (14-22, 27-36, 45-47, 50-58, 81) in addition
// to scanning format strings for date/time letters. Quoted literals, bracketed
// sections and escaped characters are ignored.
inline bool isDateFormatCode(std::string_view format, int numFmtId = -1) noexcept {
    if (numFmtId >= 0) {
        if (numFmtId == 14 || numFmtId == 15 || numFmtId == 16 || numFmtId == 17 ||
            numFmtId == 18 || numFmtId == 19 || numFmtId == 20 || numFmtId == 21 ||
            numFmtId == 22) return true;
        if (numFmtId >= 27 && numFmtId <= 36) return true;
        if (numFmtId >= 45 && numFmtId <= 47) return true;
        if (numFmtId >= 50 && numFmtId <= 58) return true;
        if (numFmtId == 81) return true;
    }
    bool quoted = false;
    for (std::size_t i = 0; i < format.size(); ++i) {
        const char c = format[i];
        if (quoted) {
            if (c == '"') quoted = false;
            continue;
        }
        if (c == '"') { quoted = true; continue; }
        if (c == '[') {
            while (i < format.size() && format[i] != ']') ++i;
            continue;
        }
        if (c == '\\') { ++i; continue; }
        if (c == 'y' || c == 'm' || c == 'd' || c == 'h' || c == 's') return true;
    }
    return false;
}

inline std::string toIso8601Date(const DateTime& value) {
    char buffer[32];
    std::snprintf(buffer, sizeof(buffer), "%04d-%02d-%02d", value.year, value.month, value.day);
    return buffer;
}

inline std::string toIso8601(const DateTime& value) {
    char buffer[64];
    const long long seconds = static_cast<long long>(value.second);
    if (value.second - static_cast<double>(seconds) <= 0.0) {
        std::snprintf(buffer, sizeof(buffer), "%04d-%02d-%02dT%02d:%02d:%02d",
                      value.year, value.month, value.day, value.hour, value.minute,
                      static_cast<int>(seconds));
    } else {
        std::snprintf(buffer, sizeof(buffer), "%04d-%02d-%02dT%02d:%02d:%06.3f",
                      value.year, value.month, value.day, value.hour, value.minute, value.second);
    }
    return buffer;
}

// Parses an ISO-8601 date, optionally with time and timezone:
//   YYYY-MM-DD [T ]HH:MM[:SS[.fff]][ Z|+HH[:MM]|-HH[:MM]]
// Offsets are applied so the result is always UTC wall-clock time.
inline std::optional<DateTime> parseIso8601(std::string_view text) {
    const auto isDigit = [](char c) { return c >= '0' && c <= '9'; };
    std::size_t i = 0;
    auto readInt = [&](int& out, int digits) -> bool {
        int value = 0;
        for (int k = 0; k < digits; ++k) {
            if (i >= text.size() || !isDigit(text[i])) return false;
            value = value * 10 + (text[i] - '0');
            ++i;
        }
        out = value;
        return true;
    };

    DateTime result;
    int year = 0, month = 0, day = 0;
    if (!readInt(year, 4)) return std::nullopt;
    if (i >= text.size() || text[i] != '-') return std::nullopt;
    ++i;
    if (!readInt(month, 2)) return std::nullopt;
    if (i >= text.size() || text[i] != '-') return std::nullopt;
    ++i;
    if (!readInt(day, 2)) return std::nullopt;
    if (month < 1 || month > 12 || day < 1 || day > 31) return std::nullopt;
    result.year = year;
    result.month = month;
    result.day = day;

    if (i < text.size()) {
        if (text[i] != 'T' && text[i] != ' ') return std::nullopt;
        ++i;
        int hour = 0, minute = 0, second = 0;
        if (!readInt(hour, 2)) return std::nullopt;
        if (i >= text.size() || text[i] != ':') return std::nullopt;
        ++i;
        if (!readInt(minute, 2)) return std::nullopt;
        if (i < text.size() && text[i] == ':') {
            ++i;
            if (!readInt(second, 2)) return std::nullopt;
        }
        double fraction = 0.0;
        if (i < text.size() && text[i] == '.') {
            ++i;
            double factor = 0.1;
            int digits = 0;
            while (i < text.size() && isDigit(text[i]) && digits < 6) {
                fraction += (text[i] - '0') * factor;
                factor *= 0.1;
                ++i;
                ++digits;
            }
        }
        if (hour > 23 || minute > 59 || second > 60) return std::nullopt;
        if (second == 60) second = 59;
        result.hour = hour;
        result.minute = minute;
        result.second = static_cast<double>(second) + fraction;

        if (i < text.size()) {
            int offsetMinutes = 0;
            if (text[i] == 'Z') {
                ++i;
            } else if (text[i] == '+' || text[i] == '-') {
                const int sign = text[i] == '-' ? -1 : 1;
                ++i;
                int offsetHour = 0, offsetMinute = 0;
                if (!readInt(offsetHour, 2)) return std::nullopt;
                if (i < text.size() && text[i] == ':') ++i;
                if (i < text.size() && isDigit(text[i])) {
                    if (!readInt(offsetMinute, 2)) return std::nullopt;
                }
                offsetMinutes = sign * (offsetHour * 60 + offsetMinute);
            } else {
                return std::nullopt;
            }
            // Local wall-clock -> UTC by shifting the serial representation.
            const double serial = toExcelSerial(result, false) - offsetMinutes / 1440.0;
            result = fromExcelSerial(serial, false);
        }
    }

    while (i < text.size() && (text[i] == ' ' || text[i] == '\t')) ++i;
    if (i != text.size()) return std::nullopt;
    return result;
}

} // namespace xlpp
