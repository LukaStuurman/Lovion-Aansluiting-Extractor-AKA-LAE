#define NOMINMAX
#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <winhttp.h>

#include <algorithm>
#include <atomic>
#include <chrono>
#include <cmath>
#include <cctype>
#include <cstdlib>
#include <ctime>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <limits>
#include <locale>
#include <map>
#include <mutex>
#include <regex>
#include <set>
#include <sstream>
#include <stdexcept>
#include <string>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#pragma comment(lib, "winhttp.lib")

namespace fs = std::filesystem;

namespace {

const char* kRdNewPrj =
    "PROJCS[\"Amersfoort / RD New\",GEOGCS[\"Amersfoort\",DATUM[\"Amersfoort\","
    "SPHEROID[\"Bessel 1841\",6377397.155,299.1528128]],PRIMEM[\"Greenwich\",0],"
    "UNIT[\"degree\",0.0174532925199433]],PROJECTION[\"Oblique_Stereographic\"],"
    "PARAMETER[\"latitude_of_origin\",52.15616055555555],"
    "PARAMETER[\"central_meridian\",5.38763888888889],"
    "PARAMETER[\"scale_factor\",0.9999079],PARAMETER[\"false_easting\",155000],"
    "PARAMETER[\"false_northing\",463000],UNIT[\"metre\",1]]";

struct Row {
    std::vector<std::string> columns;
    std::unordered_map<std::string, std::string> values;

    bool has(const std::string& name) const {
        return values.find(name) != values.end();
    }

    std::string get(const std::string& name) const {
        auto it = values.find(name);
        if (it == values.end()) {
            return {};
        }
        return it->second;
    }

    void set(const std::string& name, const std::string& value) {
        if (!has(name)) {
            columns.push_back(name);
        }
        values[name] = value;
    }
};

struct Options {
    std::string inputTextPath;
    std::string inputCsvPath;
    std::string inputExcelPath;
    std::string worksheetName = "Alle stations nieuw";
    std::string outputDirectory;
    std::string wfsUrl = "https://api.enexis.nl/opendata-assets/v1/wfs";
    std::string wfsLayerName = "Opendata:asm_e_lv_service_connection";
    std::string wfsGeoJsonPath;
    int geocodeDelayMs = 0;
    int geocodeConcurrency = 6;
    int snapDistanceMeters = 25;
    int streetFurnitureSnapDistanceMeters = 100;
    int wfsBufferMeters = 100;
    int maxRows = 0;
    bool skipGeocoding = false;
    bool skipWfs = false;
    bool help = false;
};

struct Point {
    double x = 0.0;
    double y = 0.0;
    bool valid = false;
};

struct Wgs84 {
    double lon = 0.0;
    double lat = 0.0;
};

std::string trim(const std::string& value) {
    size_t begin = 0;
    while (begin < value.size() && std::isspace(static_cast<unsigned char>(value[begin]))) {
        ++begin;
    }
    size_t end = value.size();
    while (end > begin && std::isspace(static_cast<unsigned char>(value[end - 1]))) {
        --end;
    }
    return value.substr(begin, end - begin);
}

std::string collapseWhitespace(const std::string& value) {
    std::string out;
    bool inSpace = false;
    for (unsigned char ch : value) {
        if (std::isspace(ch)) {
            if (!inSpace && !out.empty()) {
                out.push_back(' ');
            }
            inSpace = true;
        } else {
            out.push_back(static_cast<char>(ch));
            inSpace = false;
        }
    }
    if (!out.empty() && out.back() == ' ') {
        out.pop_back();
    }
    return out;
}

std::string toUpperAscii(std::string value) {
    std::transform(value.begin(), value.end(), value.begin(), [](unsigned char ch) {
        return static_cast<char>(std::toupper(ch));
    });
    return value;
}

std::string toLowerAscii(std::string value) {
    std::transform(value.begin(), value.end(), value.begin(), [](unsigned char ch) {
        return static_cast<char>(std::tolower(ch));
    });
    return value;
}

bool iequals(const std::string& a, const std::string& b) {
    return toLowerAscii(a) == toLowerAscii(b);
}

bool icontains(const std::string& haystack, const std::string& needle) {
    return toLowerAscii(haystack).find(toLowerAscii(needle)) != std::string::npos;
}

bool startsWithIcase(const std::string& value, const std::string& prefix) {
    if (value.size() < prefix.size()) {
        return false;
    }
    return iequals(value.substr(0, prefix.size()), prefix);
}

std::string normalizeText(const std::string& value) {
    return toUpperAscii(collapseWhitespace(trim(value)));
}

std::string normalizePostcode(const std::string& value) {
    std::string out;
    for (unsigned char ch : value) {
        if (!std::isspace(ch)) {
            out.push_back(static_cast<char>(std::toupper(ch)));
        }
    }
    return out;
}

std::string normalizeHouseNumber(const std::string& value) {
    std::string out;
    for (unsigned char ch : value) {
        if (!std::isspace(ch)) {
            out.push_back(static_cast<char>(std::toupper(ch)));
        }
    }
    return out;
}

std::string houseNumberDigits(const std::string& value) {
    std::string normalized = normalizeHouseNumber(value);
    if (normalized.empty()) {
        return {};
    }
    size_t index = 0;
    while (index < normalized.size() && std::isdigit(static_cast<unsigned char>(normalized[index]))) {
        ++index;
    }
    if (index == 0) {
        return normalized;
    }
    return normalized.substr(0, index);
}

std::string normalizeComparableText(const std::string& value) {
    std::string normalized = normalizeText(value);
    std::string out;
    for (unsigned char ch : normalized) {
        if (std::isalnum(ch)) {
            out.push_back(static_cast<char>(ch));
        }
    }
    return out;
}

std::string joinNonEmpty(const std::vector<std::string>& parts) {
    std::vector<std::string> kept;
    for (const auto& part : parts) {
        std::string t = trim(part);
        if (!t.empty()) {
            kept.push_back(t);
        }
    }
    std::ostringstream out;
    for (size_t i = 0; i < kept.size(); ++i) {
        if (i != 0) {
            out << ' ';
        }
        out << kept[i];
    }
    return out.str();
}

std::string replaceAll(std::string value, const std::string& from, const std::string& to) {
    size_t pos = 0;
    while ((pos = value.find(from, pos)) != std::string::npos) {
        value.replace(pos, from.size(), to);
        pos += to.size();
    }
    return value;
}

std::wstring utf8ToWide(const std::string& value) {
    if (value.empty()) {
        return {};
    }
    int size = MultiByteToWideChar(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), nullptr, 0);
    if (size <= 0) {
        return {};
    }
    std::wstring out(static_cast<size_t>(size), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), out.data(), size);
    return out;
}

std::string wideToUtf8(const std::wstring& value) {
    if (value.empty()) {
        return {};
    }
    int size = WideCharToMultiByte(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), nullptr, 0, nullptr, nullptr);
    if (size <= 0) {
        return {};
    }
    std::string out(static_cast<size_t>(size), '\0');
    WideCharToMultiByte(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), out.data(), size, nullptr, nullptr);
    return out;
}

bool isValidUtf8(const std::string& bytes) {
    if (bytes.empty()) {
        return true;
    }
    return MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, bytes.data(), static_cast<int>(bytes.size()), nullptr, 0) > 0;
}

std::string bytesToUtf8(const std::string& bytes) {
    if (bytes.size() >= 3 &&
        static_cast<unsigned char>(bytes[0]) == 0xEF &&
        static_cast<unsigned char>(bytes[1]) == 0xBB &&
        static_cast<unsigned char>(bytes[2]) == 0xBF) {
        return bytes.substr(3);
    }
    if (isValidUtf8(bytes)) {
        return bytes;
    }
    int wideSize = MultiByteToWideChar(CP_ACP, 0, bytes.data(), static_cast<int>(bytes.size()), nullptr, 0);
    if (wideSize <= 0) {
        return bytes;
    }
    std::wstring wide(static_cast<size_t>(wideSize), L'\0');
    MultiByteToWideChar(CP_ACP, 0, bytes.data(), static_cast<int>(bytes.size()), wide.data(), wideSize);
    return wideToUtf8(wide);
}

std::string toCp1252(const std::string& utf8) {
    std::wstring wide = utf8ToWide(utf8);
    if (wide.empty() && !utf8.empty()) {
        return std::string(utf8.size(), '?');
    }
    int size = WideCharToMultiByte(1252, 0, wide.data(), static_cast<int>(wide.size()), nullptr, 0, "?", nullptr);
    std::string out(static_cast<size_t>(std::max(size, 0)), '\0');
    if (size > 0) {
        WideCharToMultiByte(1252, 0, wide.data(), static_cast<int>(wide.size()), out.data(), size, "?", nullptr);
    }
    return out;
}

std::string readFileUtf8(const fs::path& path) {
    std::ifstream in(path, std::ios::binary);
    if (!in) {
        throw std::runtime_error("Kan bestand niet lezen: " + path.string());
    }
    std::ostringstream buffer;
    buffer << in.rdbuf();
    return bytesToUtf8(buffer.str());
}

void writeTextUtf8(const fs::path& path, const std::string& text, bool withBom = false) {
    std::ofstream out(path, std::ios::binary);
    if (!out) {
        throw std::runtime_error("Kan bestand niet schrijven: " + path.string());
    }
    if (withBom) {
        const unsigned char bom[] = {0xEF, 0xBB, 0xBF};
        out.write(reinterpret_cast<const char*>(bom), 3);
    }
    out.write(text.data(), static_cast<std::streamsize>(text.size()));
}

bool parseDoubleInvariant(const std::string& value, double& out) {
    std::string text = trim(value);
    if (text.empty()) {
        return false;
    }
    std::replace(text.begin(), text.end(), ',', '.');
    char* end = nullptr;
    out = std::strtod(text.c_str(), &end);
    return end != text.c_str() && trim(end ? std::string(end) : std::string()).empty();
}

std::string formatDouble(double value, int decimals) {
    std::ostringstream out;
    out.imbue(std::locale::classic());
    out << std::fixed << std::setprecision(decimals) << value;
    return out.str();
}

std::string formatDoubleCompact(double value) {
    std::ostringstream out;
    out.imbue(std::locale::classic());
    out << std::setprecision(12) << value;
    std::string text = out.str();
    if (text.find('.') != std::string::npos) {
        while (!text.empty() && text.back() == '0') {
            text.pop_back();
        }
        if (!text.empty() && text.back() == '.') {
            text.pop_back();
        }
    }
    return text;
}

std::string boolText(bool value) {
    return value ? "True" : "False";
}

std::tm localTime(std::time_t value) {
    std::tm result{};
    std::tm* local = std::localtime(&value);
    if (local) {
        result = *local;
    }
    return result;
}

Wgs84 rdToWgs84(double x, double y) {
    const double dX = (x - 155000.0) * 0.00001;
    const double dY = (y - 463000.0) * 0.00001;

    const double latSeconds =
        (3235.65389 * dY) +
        (-32.58297 * std::pow(dX, 2)) +
        (-0.2475 * std::pow(dY, 2)) +
        (-0.84978 * std::pow(dX, 2) * dY) +
        (-0.0655 * std::pow(dY, 3)) +
        (-0.01709 * std::pow(dX, 2) * std::pow(dY, 2)) +
        (-0.00738 * dX) +
        (0.0053 * std::pow(dX, 4)) +
        (-0.00039 * std::pow(dX, 2) * std::pow(dY, 3)) +
        (0.00033 * std::pow(dX, 4) * dY) +
        (-0.00012 * dX * dY);

    const double lonSeconds =
        (5260.52916 * dX) +
        (105.94684 * dX * dY) +
        (2.45656 * dX * std::pow(dY, 2)) +
        (-0.81885 * std::pow(dX, 3)) +
        (0.05594 * dX * std::pow(dY, 3)) +
        (-0.05607 * std::pow(dX, 3) * dY) +
        (0.01199 * dY) +
        (-0.00256 * std::pow(dX, 3) * std::pow(dY, 2)) +
        (0.00128 * dX * std::pow(dY, 4)) +
        (0.00022 * std::pow(dY, 2)) +
        (-0.00022 * std::pow(dX, 2)) +
        (0.00026 * std::pow(dX, 5));

    return {
        5.38720621 + (lonSeconds / 3600.0),
        52.15517440 + (latSeconds / 3600.0)
    };
}

struct Json {
    enum class Type { Null, Bool, Number, String, Array, Object };
    Type type = Type::Null;
    bool boolean = false;
    double number = 0.0;
    std::string string;
    std::vector<Json> array;
    std::map<std::string, Json> object;

    bool isNull() const { return type == Type::Null; }
    bool isObject() const { return type == Type::Object; }
    bool isArray() const { return type == Type::Array; }
    bool isString() const { return type == Type::String; }
    bool isNumber() const { return type == Type::Number; }
    bool isBool() const { return type == Type::Bool; }
};

void appendUtf8Codepoint(std::string& out, unsigned codepoint) {
    if (codepoint <= 0x7F) {
        out.push_back(static_cast<char>(codepoint));
    } else if (codepoint <= 0x7FF) {
        out.push_back(static_cast<char>(0xC0 | ((codepoint >> 6) & 0x1F)));
        out.push_back(static_cast<char>(0x80 | (codepoint & 0x3F)));
    } else if (codepoint <= 0xFFFF) {
        out.push_back(static_cast<char>(0xE0 | ((codepoint >> 12) & 0x0F)));
        out.push_back(static_cast<char>(0x80 | ((codepoint >> 6) & 0x3F)));
        out.push_back(static_cast<char>(0x80 | (codepoint & 0x3F)));
    } else {
        out.push_back(static_cast<char>(0xF0 | ((codepoint >> 18) & 0x07)));
        out.push_back(static_cast<char>(0x80 | ((codepoint >> 12) & 0x3F)));
        out.push_back(static_cast<char>(0x80 | ((codepoint >> 6) & 0x3F)));
        out.push_back(static_cast<char>(0x80 | (codepoint & 0x3F)));
    }
}

class JsonParser {
public:
    explicit JsonParser(const std::string& text) : text_(text) {}

    Json parse() {
        Json value = parseValue();
        skipWhitespace();
        if (pos_ != text_.size()) {
            throw std::runtime_error("Ongeldige JSON: extra data");
        }
        return value;
    }

private:
    const std::string& text_;
    size_t pos_ = 0;

    void skipWhitespace() {
        while (pos_ < text_.size() && std::isspace(static_cast<unsigned char>(text_[pos_]))) {
            ++pos_;
        }
    }

    char peek() const {
        if (pos_ >= text_.size()) {
            return '\0';
        }
        return text_[pos_];
    }

    char get() {
        if (pos_ >= text_.size()) {
            throw std::runtime_error("Ongeldige JSON: onverwacht einde");
        }
        return text_[pos_++];
    }

    bool consume(char expected) {
        skipWhitespace();
        if (peek() == expected) {
            ++pos_;
            return true;
        }
        return false;
    }

    Json parseValue() {
        skipWhitespace();
        char ch = peek();
        if (ch == '"') {
            Json value;
            value.type = Json::Type::String;
            value.string = parseString();
            return value;
        }
        if (ch == '{') {
            return parseObject();
        }
        if (ch == '[') {
            return parseArray();
        }
        if (startsWith("true")) {
            pos_ += 4;
            Json value;
            value.type = Json::Type::Bool;
            value.boolean = true;
            return value;
        }
        if (startsWith("false")) {
            pos_ += 5;
            Json value;
            value.type = Json::Type::Bool;
            value.boolean = false;
            return value;
        }
        if (startsWith("null")) {
            pos_ += 4;
            return Json{};
        }
        return parseNumber();
    }

    bool startsWith(const char* literal) const {
        size_t len = std::strlen(literal);
        return text_.compare(pos_, len, literal) == 0;
    }

    int hexValue(char ch) {
        if (ch >= '0' && ch <= '9') return ch - '0';
        if (ch >= 'a' && ch <= 'f') return 10 + ch - 'a';
        if (ch >= 'A' && ch <= 'F') return 10 + ch - 'A';
        return -1;
    }

    unsigned parseHex4() {
        unsigned value = 0;
        for (int i = 0; i < 4; ++i) {
            int digit = hexValue(get());
            if (digit < 0) {
                throw std::runtime_error("Ongeldige JSON unicode escape");
            }
            value = (value << 4) | static_cast<unsigned>(digit);
        }
        return value;
    }

    std::string parseString() {
        if (get() != '"') {
            throw std::runtime_error("Ongeldige JSON string");
        }
        std::string out;
        while (true) {
            char ch = get();
            if (ch == '"') {
                break;
            }
            if (ch == '\\') {
                char esc = get();
                switch (esc) {
                    case '"': out.push_back('"'); break;
                    case '\\': out.push_back('\\'); break;
                    case '/': out.push_back('/'); break;
                    case 'b': out.push_back('\b'); break;
                    case 'f': out.push_back('\f'); break;
                    case 'n': out.push_back('\n'); break;
                    case 'r': out.push_back('\r'); break;
                    case 't': out.push_back('\t'); break;
                    case 'u': {
                        unsigned cp = parseHex4();
                        if (cp >= 0xD800 && cp <= 0xDBFF && pos_ + 6 <= text_.size() && text_[pos_] == '\\' && text_[pos_ + 1] == 'u') {
                            pos_ += 2;
                            unsigned low = parseHex4();
                            if (low >= 0xDC00 && low <= 0xDFFF) {
                                cp = 0x10000 + (((cp - 0xD800) << 10) | (low - 0xDC00));
                            }
                        }
                        appendUtf8Codepoint(out, cp);
                        break;
                    }
                    default:
                        throw std::runtime_error("Ongeldige JSON escape");
                }
            } else {
                out.push_back(ch);
            }
        }
        return out;
    }

    Json parseNumber() {
        size_t begin = pos_;
        if (peek() == '-') ++pos_;
        while (std::isdigit(static_cast<unsigned char>(peek()))) ++pos_;
        if (peek() == '.') {
            ++pos_;
            while (std::isdigit(static_cast<unsigned char>(peek()))) ++pos_;
        }
        if (peek() == 'e' || peek() == 'E') {
            ++pos_;
            if (peek() == '+' || peek() == '-') ++pos_;
            while (std::isdigit(static_cast<unsigned char>(peek()))) ++pos_;
        }
        if (begin == pos_) {
            throw std::runtime_error("Ongeldige JSON waarde");
        }
        Json value;
        value.type = Json::Type::Number;
        value.number = std::strtod(text_.substr(begin, pos_ - begin).c_str(), nullptr);
        return value;
    }

    Json parseArray() {
        Json value;
        value.type = Json::Type::Array;
        get();
        skipWhitespace();
        if (consume(']')) {
            return value;
        }
        while (true) {
            value.array.push_back(parseValue());
            skipWhitespace();
            if (consume(']')) {
                break;
            }
            if (!consume(',')) {
                throw std::runtime_error("Ongeldige JSON array");
            }
        }
        return value;
    }

    Json parseObject() {
        Json value;
        value.type = Json::Type::Object;
        get();
        skipWhitespace();
        if (consume('}')) {
            return value;
        }
        while (true) {
            skipWhitespace();
            std::string key = parseString();
            if (!consume(':')) {
                throw std::runtime_error("Ongeldige JSON object");
            }
            value.object.emplace(std::move(key), parseValue());
            skipWhitespace();
            if (consume('}')) {
                break;
            }
            if (!consume(',')) {
                throw std::runtime_error("Ongeldige JSON object");
            }
        }
        return value;
    }
};

Json parseJson(const std::string& text) {
    return JsonParser(text).parse();
}

const Json* objectValue(const Json& object, const std::string& key) {
    if (!object.isObject()) {
        return nullptr;
    }
    auto it = object.object.find(key);
    if (it == object.object.end()) {
        return nullptr;
    }
    return &it->second;
}

std::string jsonValueToString(const Json* value) {
    if (value == nullptr || value->isNull()) {
        return {};
    }
    if (value->isString()) {
        return value->string;
    }
    if (value->isNumber()) {
        return formatDoubleCompact(value->number);
    }
    if (value->isBool()) {
        return boolText(value->boolean);
    }
    return {};
}

double jsonValueToDouble(const Json* value, double fallback = 0.0) {
    if (value == nullptr || value->isNull()) {
        return fallback;
    }
    if (value->isNumber()) {
        return value->number;
    }
    double parsed = 0.0;
    if (value->isString() && parseDoubleInvariant(value->string, parsed)) {
        return parsed;
    }
    return fallback;
}

const Json* findProperty(const Json& object, const std::vector<std::string>& candidates) {
    if (!object.isObject()) {
        return nullptr;
    }
    for (const auto& candidate : candidates) {
        for (const auto& entry : object.object) {
            if (iequals(entry.first, candidate)) {
                return &entry.second;
            }
        }
    }
    for (const auto& candidate : candidates) {
        for (const auto& entry : object.object) {
            if (icontains(entry.first, candidate)) {
                return &entry.second;
            }
        }
    }
    return nullptr;
}

std::string urlEncode(const std::string& value) {
    static const char* hex = "0123456789ABCDEF";
    std::string out;
    for (unsigned char ch : value) {
        if ((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') ||
            (ch >= '0' && ch <= '9') || ch == '-' || ch == '_' || ch == '.' || ch == '~') {
            out.push_back(static_cast<char>(ch));
        } else {
            out.push_back('%');
            out.push_back(hex[ch >> 4]);
            out.push_back(hex[ch & 0x0F]);
        }
    }
    return out;
}

std::string httpGet(const std::string& url) {
    std::wstring wideUrl = utf8ToWide(url);
    URL_COMPONENTS components{};
    components.dwStructSize = sizeof(components);
    components.dwSchemeLength = static_cast<DWORD>(-1);
    components.dwHostNameLength = static_cast<DWORD>(-1);
    components.dwUrlPathLength = static_cast<DWORD>(-1);
    components.dwExtraInfoLength = static_cast<DWORD>(-1);

    if (!WinHttpCrackUrl(wideUrl.c_str(), 0, 0, &components)) {
        throw std::runtime_error("Ongeldige URL: " + url);
    }

    std::wstring host(components.lpszHostName, components.dwHostNameLength);
    std::wstring path = L"/";
    if (components.dwUrlPathLength > 0 && components.lpszUrlPath) {
        path.assign(components.lpszUrlPath, components.dwUrlPathLength);
    }
    if (components.dwExtraInfoLength > 0 && components.lpszExtraInfo) {
        path.append(components.lpszExtraInfo, components.dwExtraInfoLength);
    }
    bool secure = components.nScheme == INTERNET_SCHEME_HTTPS;

    HINTERNET session = WinHttpOpen(L"LovionCoordinateToolCpp/1.0",
                                    WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                                    WINHTTP_NO_PROXY_NAME,
                                    WINHTTP_NO_PROXY_BYPASS,
                                    0);
    if (!session) {
        throw std::runtime_error("WinHTTP sessie kon niet worden geopend");
    }
    WinHttpSetTimeouts(session, 30000, 30000, 30000, 30000);

    HINTERNET connect = WinHttpConnect(session, host.c_str(), components.nPort, 0);
    if (!connect) {
        WinHttpCloseHandle(session);
        throw std::runtime_error("WinHTTP connectie mislukt");
    }

    HINTERNET request = WinHttpOpenRequest(connect, L"GET", path.c_str(), nullptr,
                                           WINHTTP_NO_REFERER,
                                           WINHTTP_DEFAULT_ACCEPT_TYPES,
                                           secure ? WINHTTP_FLAG_SECURE : 0);
    if (!request) {
        WinHttpCloseHandle(connect);
        WinHttpCloseHandle(session);
        throw std::runtime_error("WinHTTP request kon niet worden geopend");
    }

    DWORD redirectPolicy = WINHTTP_OPTION_REDIRECT_POLICY_ALWAYS;
    WinHttpSetOption(request, WINHTTP_OPTION_REDIRECT_POLICY, &redirectPolicy, sizeof(redirectPolicy));
    std::wstring headers = L"Accept: application/json\r\n";

    bool ok = WinHttpSendRequest(request, headers.c_str(), static_cast<DWORD>(headers.size()),
                                 WINHTTP_NO_REQUEST_DATA, 0, 0, 0) &&
              WinHttpReceiveResponse(request, nullptr);
    if (!ok) {
        WinHttpCloseHandle(request);
        WinHttpCloseHandle(connect);
        WinHttpCloseHandle(session);
        throw std::runtime_error("HTTP request mislukt");
    }

    DWORD status = 0;
    DWORD statusSize = sizeof(status);
    WinHttpQueryHeaders(request, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                        WINHTTP_HEADER_NAME_BY_INDEX, &status, &statusSize, WINHTTP_NO_HEADER_INDEX);

    std::string body;
    DWORD available = 0;
    do {
        available = 0;
        if (!WinHttpQueryDataAvailable(request, &available)) {
            break;
        }
        if (available == 0) {
            break;
        }
        std::string chunk(available, '\0');
        DWORD read = 0;
        if (!WinHttpReadData(request, chunk.data(), available, &read)) {
            break;
        }
        chunk.resize(read);
        body += chunk;
    } while (available > 0);

    WinHttpCloseHandle(request);
    WinHttpCloseHandle(connect);
    WinHttpCloseHandle(session);

    if (status < 200 || status >= 300) {
        throw std::runtime_error("HTTP status " + std::to_string(status) + " voor " + url);
    }
    return bytesToUtf8(body);
}

std::string escapeSolrPhrase(const std::string& value) {
    std::string escaped = replaceAll(value, "\\", "\\\\");
    escaped = replaceAll(escaped, "\"", "\\\"");
    return "\"" + escaped + "\"";
}

std::string buildAddressQuery(const Row& row) {
    std::string structured = joinNonEmpty({
        row.get("Straatnaam"),
        row.get("Huisnummer"),
        row.get("Postcode"),
        row.get("Woonplaats")
    });
    if (!structured.empty()) {
        return structured;
    }
    return joinNonEmpty({row.get("Adres"), row.get("Gemeente"), row.get("Woonplaats")});
}

std::string buildStreetKeyFromValues(const std::string& street, const std::string& postcode, const std::string& city) {
    std::string normalizedStreet = normalizeText(street);
    if (normalizedStreet.empty()) {
        return {};
    }
    return normalizedStreet + "|" + normalizePostcode(postcode) + "|" + normalizeText(city);
}

std::string buildAddressKey(const Row& row) {
    std::string street = normalizeText(row.get("Straatnaam"));
    std::string house = normalizeHouseNumber(row.get("Huisnummer"));
    std::string postcode = normalizePostcode(row.get("Postcode"));
    std::string city = normalizeText(row.get("Woonplaats"));
    if (street.empty() || house.empty() || postcode.empty()) {
        return {};
    }
    return street + "|" + house + "|" + postcode + "|" + city;
}

bool isStreetFurniture(const Row& row) {
    std::string usage = row.get("Gebruiksdoel");
    if (trim(usage).empty()) {
        return false;
    }
    static const std::vector<std::regex> patterns = {
        std::regex("(^|\\W)(riool|fontein|putkast)", std::regex_constants::ECMAScript | std::regex_constants::icase),
        std::regex("(^|\\W)(ov[\\- ]?kast|vri[\\- ]?kast|parkeer|slagboom|verkeer|brug|sluis)", std::regex_constants::ECMAScript | std::regex_constants::icase),
        std::regex("(^|\\W)(abri|bushalte|reclame|cai|telecom|telefoon|sirene|camera|flits|bord)", std::regex_constants::ECMAScript | std::regex_constants::icase),
        std::regex("(^|\\W)(verlichting|mast|lantaarn|marktkast|kermis|feest|afval|vuil|container|toilet|urinoir)", std::regex_constants::ECMAScript | std::regex_constants::icase),
        std::regex("(^|\\W)(bouwaansluiting)", std::regex_constants::ECMAScript | std::regex_constants::icase)
    };
    for (const auto& pattern : patterns) {
        if (std::regex_search(usage, pattern)) {
            return true;
        }
    }
    return false;
}

std::string buildPdokStructuredQuery(const Row& row) {
    std::string street = row.get("Straatnaam");
    std::string postcode = normalizePostcode(row.get("Postcode"));
    std::string city = row.get("Woonplaats");
    std::string houseRaw = row.get("Huisnummer");
    std::string houseDigits = houseNumberDigits(houseRaw);
    std::string houseFull = normalizeHouseNumber(houseRaw);
    std::string suffix;
    if (!houseFull.empty() && !houseDigits.empty() && houseFull.size() > houseDigits.size()) {
        suffix = houseFull.substr(houseDigits.size());
    }

    std::vector<std::string> clauses;
    if (!trim(street).empty()) {
        clauses.push_back("straatnaam:" + escapeSolrPhrase(trim(street)));
    }
    if (!houseDigits.empty()) {
        clauses.push_back("huisnummer:" + houseDigits);
    }
    if (!suffix.empty()) {
        if (suffix.size() == 1 && suffix[0] >= 'A' && suffix[0] <= 'Z') {
            clauses.push_back("huisletter:" + suffix);
        } else if (std::all_of(suffix.begin(), suffix.end(), [](unsigned char ch) { return std::isdigit(ch); }) && suffix.size() <= 4) {
            clauses.push_back("huisnummertoevoeging:" + escapeSolrPhrase(suffix));
        }
    }
    if (!postcode.empty()) {
        clauses.push_back("postcode:" + postcode);
    }
    if (!trim(city).empty()) {
        clauses.push_back("woonplaatsnaam:" + escapeSolrPhrase(trim(city)));
    }

    if (clauses.empty()) {
        return buildAddressQuery(row);
    }
    std::ostringstream out;
    for (size_t i = 0; i < clauses.size(); ++i) {
        if (i != 0) {
            out << " and ";
        }
        out << clauses[i];
    }
    return out.str();
}

std::string buildPdokRequestUri(const Row& row) {
    std::string query = buildPdokStructuredQuery(row);
    if (trim(query).empty()) {
        return {};
    }
    return "https://api.pdok.nl/bzk/locatieserver/search/v3_1/free?q=" + urlEncode(query) +
           "&fq=" + urlEncode("type:adres") +
           "&fq=" + urlEncode("bron:BAG") +
           "&rows=3&fl=" +
           urlEncode("id,weergavenaam,type,straatnaam,straatnaam_verkort,huisnummer,huisletter,huisnummertoevoeging,huis_nlt,postcode,woonplaatsnaam,centroide_ll,centroide_rd,score");
}

Point parseWktPoint(const std::string& wkt) {
    static const std::regex re(R"(POINT\s*\(\s*(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s*\))",
                               std::regex_constants::ECMAScript | std::regex_constants::icase);
    std::smatch match;
    if (!std::regex_search(wkt, match, re)) {
        return {};
    }
    return {std::stod(match[1].str()), std::stod(match[2].str()), true};
}

struct PdokCandidate {
    std::string id;
    std::string displayName;
    std::string type;
    double score = 0.0;
    std::string street;
    std::string streetShort;
    std::string postcode;
    std::string city;
    std::string houseNumber;
    std::string houseNumberFull;
    double rdX = 0.0;
    double rdY = 0.0;
    double lon = 0.0;
    double lat = 0.0;
    bool hasRd = false;
    bool hasLl = false;
};

PdokCandidate pdokDocToCandidate(const Json& doc) {
    PdokCandidate candidate;
    candidate.id = jsonValueToString(objectValue(doc, "id"));
    candidate.displayName = jsonValueToString(objectValue(doc, "weergavenaam"));
    candidate.type = jsonValueToString(objectValue(doc, "type"));
    candidate.score = jsonValueToDouble(objectValue(doc, "score"), 0.0);
    candidate.street = normalizeText(jsonValueToString(objectValue(doc, "straatnaam")));
    candidate.streetShort = normalizeText(jsonValueToString(objectValue(doc, "straatnaam_verkort")));
    candidate.postcode = normalizePostcode(jsonValueToString(objectValue(doc, "postcode")));
    candidate.city = normalizeText(jsonValueToString(objectValue(doc, "woonplaatsnaam")));
    candidate.houseNumber = houseNumberDigits(jsonValueToString(objectValue(doc, "huisnummer")));
    candidate.houseNumberFull = normalizeHouseNumber(jsonValueToString(objectValue(doc, "huis_nlt")));
    if (candidate.houseNumberFull.empty()) {
        candidate.houseNumberFull = normalizeHouseNumber(
            jsonValueToString(objectValue(doc, "huisnummer")) +
            jsonValueToString(objectValue(doc, "huisletter")));
    }
    Point rd = parseWktPoint(jsonValueToString(objectValue(doc, "centroide_rd")));
    Point ll = parseWktPoint(jsonValueToString(objectValue(doc, "centroide_ll")));
    if (rd.valid) {
        candidate.rdX = rd.x;
        candidate.rdY = rd.y;
        candidate.hasRd = true;
    }
    if (ll.valid) {
        candidate.lon = ll.x;
        candidate.lat = ll.y;
        candidate.hasLl = true;
    }
    return candidate;
}

struct PdokSelection {
    bool found = false;
    PdokCandidate candidate;
    bool strictMatch = false;
    std::string reason;
    double ranking = 0.0;
};

PdokSelection selectPdokCandidateForRow(const Row& row, const std::vector<Json>& docs) {
    PdokSelection best;
    if (docs.empty()) {
        return best;
    }

    std::string expectedStreet = normalizeText(row.get("Straatnaam"));
    std::string expectedPostcode = normalizePostcode(row.get("Postcode"));
    std::string expectedCity = normalizeText(row.get("Woonplaats"));
    std::string expectedHouseNumber = normalizeHouseNumber(row.get("Huisnummer"));
    std::string expectedHouseDigits = houseNumberDigits(row.get("Huisnummer"));

    for (const auto& doc : docs) {
        PdokCandidate candidate = pdokDocToCandidate(doc);
        bool streetMatches = false;
        if (!expectedStreet.empty()) {
            streetMatches = candidate.street == expectedStreet || candidate.streetShort == expectedStreet;
        }
        bool postcodeMatches = expectedPostcode.empty() || candidate.postcode == expectedPostcode;
        bool cityMatches = expectedCity.empty() || candidate.city == expectedCity;
        bool houseDigitsMatch = expectedHouseDigits.empty() || candidate.houseNumber == expectedHouseDigits;
        bool houseFullMatch = expectedHouseNumber.empty() || candidate.houseNumberFull == expectedHouseNumber;
        bool strictMatch = streetMatches && postcodeMatches && cityMatches && houseFullMatch;

        std::string reason;
        if (strictMatch) {
            reason = "exact";
        } else if (!streetMatches) {
            reason = "street_mismatch";
        } else if (!postcodeMatches) {
            reason = "postcode_mismatch";
        } else if (!houseDigitsMatch) {
            reason = "house_number_mismatch";
        } else if (!houseFullMatch) {
            reason = "house_suffix_mismatch";
        } else if (!cityMatches) {
            reason = "city_mismatch";
        } else {
            reason = "fuzzy";
        }

        double ranking =
            (strictMatch ? 1000.0 : 0.0) +
            (streetMatches ? 200.0 : 0.0) +
            (postcodeMatches ? 100.0 : 0.0) +
            (cityMatches ? 50.0 : 0.0) +
            (houseFullMatch ? 40.0 : (houseDigitsMatch ? 20.0 : 0.0)) +
            std::round(candidate.score);

        if (!best.found || ranking > best.ranking) {
            best.found = true;
            best.candidate = candidate;
            best.strictMatch = strictMatch;
            best.reason = reason;
            best.ranking = ranking;
        }
    }

    return best;
}

bool isPdokToleratedReason(const std::string& reason) {
    return trim(reason) == "house_suffix_mismatch";
}

std::vector<Json> docsFromPdokJson(const std::string& jsonText) {
    std::vector<Json> docs;
    try {
        Json root = parseJson(jsonText);
        const Json* response = objectValue(root, "response");
        const Json* docsNode = response ? objectValue(*response, "docs") : nullptr;
        if (docsNode && docsNode->isArray()) {
            docs = docsNode->array;
        }
    } catch (...) {
        docs.clear();
    }
    return docs;
}

std::unordered_map<std::string, std::vector<Json>> fetchPdokBatch(const std::vector<std::string>& queries, int concurrency, int delayMs) {
    std::unordered_map<std::string, std::vector<Json>> cache;
    if (queries.empty()) {
        return cache;
    }

    int workerCount = std::max(1, std::min<int>(concurrency, static_cast<int>(queries.size())));
    std::atomic<size_t> next{0};
    std::mutex cacheMutex;
    std::vector<std::thread> workers;

    for (int i = 0; i < workerCount; ++i) {
        workers.emplace_back([&]() {
            while (true) {
                size_t index = next.fetch_add(1);
                if (index >= queries.size()) {
                    break;
                }
                if (delayMs > 0 && index > 0) {
                    std::this_thread::sleep_for(std::chrono::milliseconds(delayMs));
                }
                std::vector<Json> docs;
                try {
                    docs = docsFromPdokJson(httpGet(queries[index]));
                } catch (...) {
                    docs.clear();
                }
                std::lock_guard<std::mutex> lock(cacheMutex);
                cache[queries[index]] = std::move(docs);
            }
        });
    }
    for (auto& worker : workers) {
        worker.join();
    }
    return cache;
}

std::vector<std::string> splitLines(const std::string& text) {
    std::vector<std::string> lines;
    std::string current;
    std::istringstream in(text);
    while (std::getline(in, current)) {
        if (!current.empty() && current.back() == '\r') {
            current.pop_back();
        }
        lines.push_back(current);
    }
    return lines;
}

std::vector<std::vector<std::string>> parseCsvText(const std::string& text, char delimiter) {
    std::vector<std::vector<std::string>> rows;
    std::vector<std::string> row;
    std::string field;
    bool inQuotes = false;

    for (size_t i = 0; i < text.size(); ++i) {
        char ch = text[i];
        if (inQuotes) {
            if (ch == '"') {
                if (i + 1 < text.size() && text[i + 1] == '"') {
                    field.push_back('"');
                    ++i;
                } else {
                    inQuotes = false;
                }
            } else {
                field.push_back(ch);
            }
        } else {
            if (ch == '"') {
                inQuotes = true;
            } else if (ch == delimiter) {
                row.push_back(field);
                field.clear();
            } else if (ch == '\n') {
                row.push_back(field);
                field.clear();
                if (!row.empty()) {
                    if (!row.back().empty() && row.back().back() == '\r') {
                        row.back().pop_back();
                    }
                    rows.push_back(row);
                }
                row.clear();
            } else {
                field.push_back(ch);
            }
        }
    }
    if (!field.empty() && field.back() == '\r') {
        field.pop_back();
    }
    row.push_back(field);
    bool any = false;
    for (const auto& item : row) {
        if (!item.empty()) {
            any = true;
            break;
        }
    }
    if (any) {
        rows.push_back(row);
    }
    return rows;
}

char detectCsvDelimiter(const std::string& text) {
    std::string firstLine;
    std::istringstream in(text);
    std::getline(in, firstLine);
    size_t comma = std::count(firstLine.begin(), firstLine.end(), ',');
    size_t semicolon = std::count(firstLine.begin(), firstLine.end(), ';');
    return semicolon > comma ? ';' : ',';
}

std::vector<Row> readCsvRows(const fs::path& path, int maxRows) {
    std::cout << "[INFO] Lees CSV: " << path.string() << "\n";
    std::string text = readFileUtf8(path);
    char delimiter = detectCsvDelimiter(text);
    auto table = parseCsvText(text, delimiter);
    if (table.size() < 2) {
        throw std::runtime_error("CSV bevat geen datarijen.");
    }

    std::vector<std::string> headers = table.front();
    std::unordered_map<std::string, int> used;
    for (size_t i = 0; i < headers.size(); ++i) {
        headers[i] = trim(headers[i]);
        if (headers[i].empty()) {
            headers[i] = "Column" + std::to_string(i + 1);
        }
        std::string base = headers[i];
        int& count = used[base];
        ++count;
        if (count > 1) {
            headers[i] = base + "_" + std::to_string(count);
        }
    }

    std::vector<Row> rows;
    for (size_t r = 1; r < table.size(); ++r) {
        if (maxRows > 0 && static_cast<int>(rows.size()) >= maxRows) {
            break;
        }
        Row row;
        row.set("_RowNumber", std::to_string(r + 1));
        bool hasValue = false;
        for (size_t c = 0; c < headers.size(); ++c) {
            std::string value = c < table[r].size() ? table[r][c] : "";
            if (!trim(value).empty()) {
                hasValue = true;
            }
            row.set(headers[c], value);
        }
        if (hasValue) {
            rows.push_back(std::move(row));
        }
    }
    return rows;
}

struct Overdrachtspunt {
    double x1 = 0.0;
    double y1 = 0.0;
    double x2 = 0.0;
    double y2 = 0.0;
    bool valid = false;
    bool isPoint = false;
};

Overdrachtspunt parseOverdrachtspuntValue(const std::string& value) {
    static const std::regex re(
        R"(linksonder:\s*(-?\d+(?:,\d+)?)\s*:\s*(-?\d+(?:,\d+)?)\s*m\s*/\s*rechtsboven:\s*(-?\d+(?:,\d+)?)\s*:\s*(-?\d+(?:,\d+)?)\s*m)",
        std::regex_constants::ECMAScript | std::regex_constants::icase);
    std::smatch match;
    if (!std::regex_search(value, match, re)) {
        return {};
    }
    Overdrachtspunt result;
    result.valid = parseDoubleInvariant(match[1].str(), result.x1) &&
                   parseDoubleInvariant(match[2].str(), result.y1) &&
                   parseDoubleInvariant(match[3].str(), result.x2) &&
                   parseDoubleInvariant(match[4].str(), result.y2);
    result.isPoint = result.valid && result.x1 == result.x2 && result.y1 == result.y2;
    return result;
}

void finalizeLovionTextRecord(const Row& record, std::vector<Row>& rows) {
    if (record.columns.empty()) {
        return;
    }
    Row row = record;
    Overdrachtspunt point = parseOverdrachtspuntValue(row.get("Overdrachtspunt"));
    if (point.valid) {
        row.set("OverdrachtspuntRdX1", formatDouble(point.x1, 3));
        row.set("OverdrachtspuntRdY1", formatDouble(point.y1, 3));
        row.set("OverdrachtspuntRdX2", formatDouble(point.x2, 3));
        row.set("OverdrachtspuntRdY2", formatDouble(point.y2, 3));
        row.set("OverdrachtspuntIsPoint", boolText(point.isPoint));
        row.set("FinalRdX", formatDouble(point.x1, 3));
        row.set("FinalRdY", formatDouble(point.y1, 3));
        Wgs84 wgs = rdToWgs84(point.x1, point.y1);
        row.set("FinalLon", formatDouble(wgs.lon, 8));
        row.set("FinalLat", formatDouble(wgs.lat, 8));
        row.set("CoordinateSource", "overdrachtspunt_text");
        row.set("MatchStatus", point.isPoint ? "overdrachtspunt_exact" : "overdrachtspunt_centroid");
    } else {
        row.set("OverdrachtspuntIsPoint", "False");
        row.set("FinalRdX", "");
        row.set("FinalRdY", "");
        row.set("FinalLon", "");
        row.set("FinalLat", "");
        row.set("CoordinateSource", "none");
        row.set("MatchStatus", "overdrachtspunt_missing");
    }
    rows.push_back(std::move(row));
}

std::vector<Row> readLovionTextRows(const fs::path& path, int maxRows) {
    std::cout << "[INFO] Lees tekstexport: " << path.string() << "\n";
    std::vector<std::string> lines = splitLines(readFileUtf8(path));
    std::vector<Row> rows;
    Row current;
    bool hasCurrent = false;
    int recordNumber = 0;
    std::regex dateRe(R"(^\((\d{1,2}-\d{1,2}-\d{4}\s+\d{2}:\d{2}:\d{2})\)$)");
    std::regex kvRe(R"(^([^:]+):\s*(.*)$)");

    for (const auto& rawLine : lines) {
        std::string line = trim(rawLine);
        if (startsWithIcase(line, "LS Aansluiting")) {
            if (hasCurrent) {
                finalizeLovionTextRecord(current, rows);
                if (maxRows > 0 && static_cast<int>(rows.size()) >= maxRows) {
                    return rows;
                }
            }
            current = Row{};
            hasCurrent = true;
            ++recordNumber;
            current.set("_RecordNumber", std::to_string(recordNumber));
            current.set("RecordTitle", line);
            continue;
        }
        if (!hasCurrent || line.empty()) {
            continue;
        }
        std::smatch match;
        if (std::regex_match(line, match, dateRe)) {
            current.set("ExportedAt", match[1].str());
            continue;
        }
        if (std::regex_match(line, match, kvRe)) {
            current.set(trim(match[1].str()), match[2].str());
            continue;
        }
        if (!current.has("ExportedBy")) {
            current.set("ExportedBy", line);
        }
    }
    if (hasCurrent && (maxRows <= 0 || static_cast<int>(rows.size()) < maxRows)) {
        finalizeLovionTextRecord(current, rows);
    }
    return rows;
}

struct BoundingBox {
    double minX = 0.0;
    double minY = 0.0;
    double maxX = 0.0;
    double maxY = 0.0;
    bool valid = false;
};

BoundingBox boundingBoxFromPdokRows(const std::vector<Row>& rows, int bufferMeters) {
    BoundingBox box;
    bool first = true;
    for (const auto& row : rows) {
        double x = 0.0;
        double y = 0.0;
        if (!parseDoubleInvariant(row.get("PdokRdX"), x) || !parseDoubleInvariant(row.get("PdokRdY"), y)) {
            continue;
        }
        if (first) {
            box.minX = box.maxX = x;
            box.minY = box.maxY = y;
            first = false;
        } else {
            box.minX = std::min(box.minX, x);
            box.maxX = std::max(box.maxX, x);
            box.minY = std::min(box.minY, y);
            box.maxY = std::max(box.maxY, y);
        }
    }
    if (first) {
        return {};
    }
    box.minX -= bufferMeters;
    box.maxX += bufferMeters;
    box.minY -= bufferMeters;
    box.maxY += bufferMeters;
    box.valid = true;
    return box;
}

std::string buildQueryString(const std::vector<std::pair<std::string, std::string>>& parameters) {
    std::ostringstream out;
    for (size_t i = 0; i < parameters.size(); ++i) {
        if (i != 0) {
            out << '&';
        }
        out << urlEncode(parameters[i].first) << '=' << urlEncode(parameters[i].second);
    }
    return out.str();
}

std::vector<Json> featuresFromGeoJsonText(const std::string& text) {
    Json root = parseJson(text);
    const Json* features = objectValue(root, "features");
    if (features && features->isArray()) {
        return features->array;
    }
    return {};
}

std::vector<Json> loadWfsFeaturesFromService(const std::string& baseUrl, const std::string& layerName, const BoundingBox& box) {
    if (!box.valid) {
        return {};
    }
    std::ostringstream bbox;
    bbox.imbue(std::locale::classic());
    bbox << box.minX << ',' << box.minY << ',' << box.maxX << ',' << box.maxY << ",EPSG:28992";

    std::string separator = baseUrl.find('?') == std::string::npos ? "?" : "&";
    std::string url = baseUrl + separator + buildQueryString({
        {"service", "WFS"},
        {"version", "2.0.0"},
        {"request", "GetFeature"},
        {"typeNames", layerName},
        {"outputFormat", "application/json"},
        {"srsName", "EPSG:28992"},
        {"bbox", bbox.str()},
        {"count", "5000"}
    });

    std::cout << "[INFO] Laad WFS-features uit bbox: " << bbox.str() << "\n";
    return featuresFromGeoJsonText(httpGet(url));
}

void collectCoordinatePairs(const Json& node, std::vector<Point>& pairs) {
    if (!node.isArray()) {
        return;
    }
    if (node.array.size() >= 2 && node.array[0].isNumber() && node.array[1].isNumber()) {
        pairs.push_back({node.array[0].number, node.array[1].number, true});
        return;
    }
    for (const auto& child : node.array) {
        collectCoordinatePairs(child, pairs);
    }
}

Point geometryCentroid(const Json& geometry) {
    const Json* coordinates = objectValue(geometry, "coordinates");
    if (!coordinates) {
        return {};
    }
    std::vector<Point> pairs;
    collectCoordinatePairs(*coordinates, pairs);
    if (pairs.empty()) {
        return {};
    }
    double sumX = 0.0;
    double sumY = 0.0;
    for (const auto& pair : pairs) {
        sumX += pair.x;
        sumY += pair.y;
    }
    return {sumX / pairs.size(), sumY / pairs.size(), true};
}

std::string buildWfsAddressKey(const Json& properties) {
    std::string street = normalizeText(jsonValueToString(findProperty(properties, {
        "straatnaam", "street", "openbareruimtenaam", "openbare_ruimte_naam", "road"
    })));
    std::string house = normalizeHouseNumber(jsonValueToString(findProperty(properties, {
        "huisnummer", "house_number", "nummer", "number", "huis_nummer"
    })));
    std::string postcode = normalizePostcode(jsonValueToString(findProperty(properties, {
        "postcode", "postalcode", "zip", "zip_code", "pc6"
    })));
    std::string city = normalizeText(jsonValueToString(findProperty(properties, {
        "woonplaats", "woonplaatsnaam", "city", "town", "place"
    })));
    if (street.empty() || house.empty() || postcode.empty()) {
        return {};
    }
    return street + "|" + house + "|" + postcode + "|" + city;
}

std::string buildWfsStreetKey(const Json& properties) {
    return buildStreetKeyFromValues(
        jsonValueToString(findProperty(properties, {"straatnaam", "street", "openbareruimtenaam", "openbare_ruimte_naam", "road"})),
        jsonValueToString(findProperty(properties, {"postcode", "postalcode", "zip", "zip_code", "pc6"})),
        jsonValueToString(findProperty(properties, {"woonplaats", "woonplaatsnaam", "city", "town", "place"}))
    );
}

struct WfsFeature {
    std::string featureId;
    std::string geometryType;
    double rdX = 0.0;
    double rdY = 0.0;
    std::string addressKey;
    std::string streetKey;
};

std::vector<WfsFeature> prepareWfsFeatures(const std::vector<Json>& rawFeatures) {
    std::vector<WfsFeature> prepared;
    for (const auto& feature : rawFeatures) {
        if (!feature.isObject()) {
            continue;
        }
        const Json* geometry = objectValue(feature, "geometry");
        const Json* properties = objectValue(feature, "properties");
        if (!geometry || !properties) {
            continue;
        }
        Point centroid = geometryCentroid(*geometry);
        if (!centroid.valid) {
            continue;
        }
        WfsFeature out;
        out.featureId = jsonValueToString(objectValue(feature, "id"));
        out.geometryType = jsonValueToString(objectValue(*geometry, "type"));
        out.rdX = centroid.x;
        out.rdY = centroid.y;
        out.addressKey = buildWfsAddressKey(*properties);
        out.streetKey = buildWfsStreetKey(*properties);
        prepared.push_back(std::move(out));
    }
    return prepared;
}

double rdDistance(double x1, double y1, double x2, double y2) {
    double dx = x2 - x1;
    double dy = y2 - y1;
    return std::sqrt(dx * dx + dy * dy);
}

struct WfsNearest {
    bool found = false;
    size_t index = 0;
    double distance = 0.0;
};

WfsNearest nearestWfsFeature(const std::vector<WfsFeature>& features, const std::vector<size_t>& indices,
                             double baseX, double baseY, double maxDistance = std::numeric_limits<double>::infinity()) {
    WfsNearest best;
    for (size_t index : indices) {
        double distance = rdDistance(baseX, baseY, features[index].rdX, features[index].rdY);
        if (!best.found || distance < best.distance) {
            best.found = true;
            best.index = index;
            best.distance = distance;
        }
    }
    if (best.found && best.distance > maxDistance) {
        return {};
    }
    return best;
}

std::vector<size_t> allFeatureIndices(const std::vector<WfsFeature>& features) {
    std::vector<size_t> indices(features.size());
    for (size_t i = 0; i < features.size(); ++i) {
        indices[i] = i;
    }
    return indices;
}

void initializeCoordinateColumns(Row& row) {
    bool streetFurniture = isStreetFurniture(row);
    row.set("AddressQuery", buildAddressQuery(row));
    row.set("AddressKey", buildAddressKey(row));
    row.set("StreetKey", buildStreetKeyFromValues(row.get("Straatnaam"), row.get("Postcode"), row.get("Woonplaats")));
    row.set("IsStreetFurniture", boolText(streetFurniture));
    row.set("UsageCategory", streetFurniture ? "street_furniture" : "service_connection");
    row.set("PdokDisplayName", "");
    row.set("PdokType", "");
    row.set("PdokScore", "");
    row.set("PdokRdX", "");
    row.set("PdokRdY", "");
    row.set("PdokLon", "");
    row.set("PdokLat", "");
    row.set("PdokStrictMatch", "False");
    row.set("PdokMatchReason", "");
    row.set("WfsFeatureId", "");
    row.set("WfsMatchMode", "");
    row.set("WfsDistanceM", "");
    row.set("FinalRdX", "");
    row.set("FinalRdY", "");
    row.set("FinalLon", "");
    row.set("FinalLat", "");
    row.set("CoordinateSource", "none");
    row.set("MatchStatus", "unmatched");
}

void geocodeRowsViaPdok(std::vector<Row>& rows, const Options& options) {
    std::vector<std::string> queries;
    std::unordered_set<std::string> seen;
    for (const auto& row : rows) {
        std::string uri = buildPdokRequestUri(row);
        if (!uri.empty() && seen.insert(uri).second) {
            queries.push_back(uri);
        }
    }

    std::cout << "[INFO] Geocodeer unieke adressen via PDOK: " << queries.size() << "\n";
    auto cache = fetchPdokBatch(queries, options.geocodeConcurrency, options.geocodeDelayMs);

    for (auto& row : rows) {
        std::string uri = buildPdokRequestUri(row);
        auto it = cache.find(uri);
        std::vector<Json> empty;
        const std::vector<Json>& docs = it == cache.end() ? empty : it->second;
        PdokSelection result = selectPdokCandidateForRow(row, docs);
        if (!result.found) {
            row.set("MatchStatus", "pdok_not_found");
            continue;
        }

        const PdokCandidate& candidate = result.candidate;
        row.set("PdokDisplayName", candidate.displayName);
        row.set("PdokType", candidate.type);
        row.set("PdokScore", formatDoubleCompact(candidate.score));
        if (candidate.hasRd) {
            row.set("PdokRdX", formatDouble(candidate.rdX, 3));
            row.set("PdokRdY", formatDouble(candidate.rdY, 3));
        }
        if (candidate.hasLl) {
            row.set("PdokLon", formatDouble(candidate.lon, 8));
            row.set("PdokLat", formatDouble(candidate.lat, 8));
        }
        row.set("PdokStrictMatch", boolText(result.strictMatch));
        row.set("PdokMatchReason", result.reason);

        bool tolerated = isPdokToleratedReason(result.reason);
        bool streetFurniture = row.get("IsStreetFurniture") == "True";
        if (streetFurniture) {
            if (result.strictMatch) {
                row.set("MatchStatus", "street_furniture_seed_exact");
            } else if (tolerated) {
                row.set("MatchStatus", "street_furniture_seed_tolerant_house_suffix");
            } else {
                row.set("MatchStatus", "street_furniture_seed_fuzzy");
            }
            continue;
        }

        if (!result.strictMatch && !tolerated) {
            row.set("MatchStatus", "pdok_not_exact");
            continue;
        }
        if (candidate.hasRd) {
            row.set("FinalRdX", formatDouble(candidate.rdX, 3));
            row.set("FinalRdY", formatDouble(candidate.rdY, 3));
        }
        if (candidate.hasLl) {
            row.set("FinalLon", formatDouble(candidate.lon, 8));
            row.set("FinalLat", formatDouble(candidate.lat, 8));
        }
        if (result.strictMatch) {
            row.set("CoordinateSource", "pdok");
            row.set("MatchStatus", "pdok_exact");
        } else {
            row.set("CoordinateSource", "pdok_tolerant");
            row.set("MatchStatus", "pdok_tolerant_house_suffix");
        }
    }
}

void matchRowsToWfs(std::vector<Row>& rows, const std::vector<WfsFeature>& features, const Options& options) {
    std::unordered_map<std::string, std::vector<size_t>> byAddress;
    std::unordered_map<std::string, std::vector<size_t>> byStreet;
    for (size_t i = 0; i < features.size(); ++i) {
        if (!features[i].addressKey.empty()) {
            byAddress[features[i].addressKey].push_back(i);
        }
        if (!features[i].streetKey.empty()) {
            byStreet[features[i].streetKey].push_back(i);
        }
    }
    std::vector<size_t> all = allFeatureIndices(features);

    for (auto& row : rows) {
        double pdokX = 0.0;
        double pdokY = 0.0;
        if (!parseDoubleInvariant(row.get("PdokRdX"), pdokX) || !parseDoubleInvariant(row.get("PdokRdY"), pdokY)) {
            continue;
        }

        WfsNearest selected;
        std::string matchMode;
        bool streetFurniture = row.get("IsStreetFurniture") == "True";
        if (streetFurniture) {
            auto streetIt = byStreet.find(row.get("StreetKey"));
            if (streetIt != byStreet.end()) {
                selected = nearestWfsFeature(features, streetIt->second, pdokX, pdokY,
                                             static_cast<double>(options.streetFurnitureSnapDistanceMeters));
                if (selected.found) {
                    matchMode = "street_furniture_street";
                }
            }
            if (!selected.found) {
                selected = nearestWfsFeature(features, all, pdokX, pdokY,
                                             static_cast<double>(options.streetFurnitureSnapDistanceMeters));
                if (selected.found) {
                    matchMode = "street_furniture_nearest";
                }
            }
            if (!selected.found) {
                row.set("FinalRdX", "");
                row.set("FinalRdY", "");
                row.set("FinalLon", "");
                row.set("FinalLat", "");
                row.set("CoordinateSource", "none");
                row.set("MatchStatus", "street_furniture_no_wfs");
                continue;
            }
        } else {
            if (row.get("PdokStrictMatch") != "True" && !isPdokToleratedReason(row.get("PdokMatchReason"))) {
                continue;
            }
            auto addressIt = byAddress.find(row.get("AddressKey"));
            if (addressIt != byAddress.end()) {
                selected = nearestWfsFeature(features, addressIt->second, pdokX, pdokY);
                if (selected.found) {
                    matchMode = "address_key";
                }
            }
            if (!selected.found) {
                selected = nearestWfsFeature(features, all, pdokX, pdokY,
                                             static_cast<double>(options.snapDistanceMeters));
                if (selected.found) {
                    matchMode = "nearest";
                }
            }
            if (!selected.found) {
                continue;
            }
        }

        const WfsFeature& feature = features[selected.index];
        row.set("WfsFeatureId", feature.featureId);
        row.set("WfsMatchMode", matchMode);
        row.set("WfsDistanceM", formatDouble(selected.distance, 2));
        row.set("FinalRdX", formatDouble(feature.rdX, 3));
        row.set("FinalRdY", formatDouble(feature.rdY, 3));
        Wgs84 wgs = rdToWgs84(feature.rdX, feature.rdY);
        row.set("FinalLon", formatDouble(wgs.lon, 8));
        row.set("FinalLat", formatDouble(wgs.lat, 8));
        row.set("CoordinateSource", "wfs");
        row.set("MatchStatus", matchMode);
    }
}

std::string csvEscape(const std::string& value) {
    return "\"" + replaceAll(value, "\"", "\"\"") + "\"";
}

std::vector<std::string> collectColumns(const std::vector<Row>& rows) {
    std::vector<std::string> columns;
    std::unordered_set<std::string> seen;
    for (const auto& row : rows) {
        for (const auto& column : row.columns) {
            if (!column.empty() && column[0] == '_') {
                continue;
            }
            if (seen.insert(column).second) {
                columns.push_back(column);
            }
        }
    }
    return columns;
}

void exportCsv(const std::vector<Row>& rows, const fs::path& path) {
    std::vector<std::string> columns = collectColumns(rows);
    std::ostringstream out;
    for (size_t i = 0; i < columns.size(); ++i) {
        if (i != 0) out << ',';
        out << csvEscape(columns[i]);
    }
    out << "\r\n";
    for (const auto& row : rows) {
        for (size_t i = 0; i < columns.size(); ++i) {
            if (i != 0) out << ',';
            out << csvEscape(row.get(columns[i]));
        }
        out << "\r\n";
    }
    writeTextUtf8(path, out.str(), true);
}

std::string jsonEscape(const std::string& value) {
    std::ostringstream out;
    for (unsigned char ch : value) {
        switch (ch) {
            case '"': out << "\\\""; break;
            case '\\': out << "\\\\"; break;
            case '\b': out << "\\b"; break;
            case '\f': out << "\\f"; break;
            case '\n': out << "\\n"; break;
            case '\r': out << "\\r"; break;
            case '\t': out << "\\t"; break;
            default:
                if (ch < 0x20) {
                    out << "\\u" << std::hex << std::setw(4) << std::setfill('0') << static_cast<int>(ch);
                } else {
                    out << static_cast<char>(ch);
                }
        }
    }
    return out.str();
}

bool isKnownNumericField(const std::string& name) {
    static const std::unordered_set<std::string> fields = {
        "OverdrachtspuntRdX1", "OverdrachtspuntRdY1", "OverdrachtspuntRdX2", "OverdrachtspuntRdY2",
        "PdokScore", "PdokRdX", "PdokRdY", "PdokLon", "PdokLat",
        "WfsDistanceM", "FinalRdX", "FinalRdY", "FinalLon", "FinalLat"
    };
    return fields.find(name) != fields.end();
}

std::string jsonPropertyValue(const std::string& name, const std::string& value) {
    if (trim(value).empty()) {
        return "null";
    }
    if (value == "True") {
        return "true";
    }
    if (value == "False") {
        return "false";
    }
    double number = 0.0;
    if (isKnownNumericField(name) && parseDoubleInvariant(value, number)) {
        return formatDoubleCompact(number);
    }
    return "\"" + jsonEscape(value) + "\"";
}

void exportGeoJson(const std::vector<Row>& rows, const fs::path& path) {
    std::vector<std::string> columns = collectColumns(rows);
    std::ostringstream out;
    out << "{\n";
    out << "  \"type\": \"FeatureCollection\",\n";
    out << "  \"name\": \"lovion_coordinates\",\n";
    out << "  \"crs\": {\"type\": \"name\", \"properties\": {\"name\": \"EPSG:28992\"}},\n";
    out << "  \"features\": [\n";
    bool firstFeature = true;
    for (const auto& row : rows) {
        double x = 0.0;
        double y = 0.0;
        if (!parseDoubleInvariant(row.get("FinalRdX"), x) || !parseDoubleInvariant(row.get("FinalRdY"), y)) {
            continue;
        }
        if (!firstFeature) {
            out << ",\n";
        }
        firstFeature = false;
        out << "    {\"type\": \"Feature\", \"geometry\": {\"type\": \"Point\", \"coordinates\": ["
            << formatDoubleCompact(x) << ", " << formatDoubleCompact(y) << "]}, \"properties\": {";
        bool firstProperty = true;
        for (const auto& column : columns) {
            if (column == "FinalRdX" || column == "FinalRdY" || column == "FinalLon" || column == "FinalLat") {
                continue;
            }
            if (!firstProperty) {
                out << ", ";
            }
            firstProperty = false;
            out << "\"" << jsonEscape(column) << "\": " << jsonPropertyValue(column, row.get(column));
        }
        out << "}}";
    }
    out << "\n  ]\n";
    out << "}\n";
    writeTextUtf8(path, out.str(), true);
}

void writeBEInt32(std::ofstream& out, int value) {
    unsigned char bytes[4] = {
        static_cast<unsigned char>((value >> 24) & 0xFF),
        static_cast<unsigned char>((value >> 16) & 0xFF),
        static_cast<unsigned char>((value >> 8) & 0xFF),
        static_cast<unsigned char>(value & 0xFF)
    };
    out.write(reinterpret_cast<const char*>(bytes), 4);
}

void writeLEInt16(std::ofstream& out, int value) {
    unsigned char bytes[2] = {
        static_cast<unsigned char>(value & 0xFF),
        static_cast<unsigned char>((value >> 8) & 0xFF)
    };
    out.write(reinterpret_cast<const char*>(bytes), 2);
}

void writeLEInt32(std::ofstream& out, int value) {
    unsigned char bytes[4] = {
        static_cast<unsigned char>(value & 0xFF),
        static_cast<unsigned char>((value >> 8) & 0xFF),
        static_cast<unsigned char>((value >> 16) & 0xFF),
        static_cast<unsigned char>((value >> 24) & 0xFF)
    };
    out.write(reinterpret_cast<const char*>(bytes), 4);
}

void writeLEDouble(std::ofstream& out, double value) {
    static_assert(sizeof(double) == 8, "double moet 8 bytes zijn");
    unsigned char bytes[8];
    std::memcpy(bytes, &value, 8);
    out.write(reinterpret_cast<const char*>(bytes), 8);
}

void writeShapefileHeader(std::ofstream& out, int fileLengthWords, int shapeType,
                          double minX, double minY, double maxX, double maxY) {
    writeBEInt32(out, 9994);
    for (int i = 0; i < 5; ++i) {
        writeBEInt32(out, 0);
    }
    writeBEInt32(out, fileLengthWords);
    writeLEInt32(out, 1000);
    writeLEInt32(out, shapeType);
    writeLEDouble(out, minX);
    writeLEDouble(out, minY);
    writeLEDouble(out, maxX);
    writeLEDouble(out, maxY);
    writeLEDouble(out, 0.0);
    writeLEDouble(out, 0.0);
    writeLEDouble(out, 0.0);
    writeLEDouble(out, 0.0);
}

struct DbfField {
    std::string name;
    std::string originalName;
    char type = 'C';
    int length = 1;
    int decimals = 0;
};

std::string dbfSafeFieldName(const std::string& name, std::unordered_set<std::string>& used) {
    std::string normalized = normalizeText(name);
    std::string ascii;
    bool previousUnderscore = false;
    for (unsigned char ch : normalized) {
        if (std::isalnum(ch)) {
            ascii.push_back(static_cast<char>(ch));
            previousUnderscore = false;
        } else if (!previousUnderscore && !ascii.empty()) {
            ascii.push_back('_');
            previousUnderscore = true;
        }
    }
    while (!ascii.empty() && ascii.back() == '_') {
        ascii.pop_back();
    }
    if (ascii.empty()) {
        ascii = "FIELD";
    }
    if (ascii.size() > 10) {
        ascii.resize(10);
    }
    if (!ascii.empty() && std::isdigit(static_cast<unsigned char>(ascii[0]))) {
        ascii = "F" + ascii.substr(0, std::min<size_t>(9, ascii.size()));
    }
    std::string candidate = ascii;
    int suffix = 1;
    while (used.find(candidate) != used.end()) {
        std::string suffixText = std::to_string(suffix++);
        size_t prefixLength = std::min<size_t>(10 - suffixText.size(), ascii.size());
        candidate = ascii.substr(0, prefixLength) + suffixText;
    }
    used.insert(candidate);
    return candidate;
}

int cp1252Length(const std::string& utf8) {
    return static_cast<int>(toCp1252(utf8).size());
}

std::vector<DbfField> buildDbfFields(const std::vector<Row>& pointRows) {
    std::vector<std::string> columns = collectColumns(pointRows);
    std::unordered_set<std::string> used;
    std::vector<DbfField> fields;
    for (const auto& column : columns) {
        DbfField field;
        field.name = dbfSafeFieldName(column, used);
        field.originalName = column;
        if (isKnownNumericField(column)) {
            field.type = 'N';
            field.length = 18;
            field.decimals = (column == "FinalLon" || column == "FinalLat" || column == "PdokLon" || column == "PdokLat") ? 8 : 3;
            if (column == "WfsDistanceM") {
                field.decimals = 2;
            }
            if (column == "PdokScore") {
                field.decimals = 2;
            }
        } else {
            field.type = 'C';
            int length = 1;
            for (const auto& row : pointRows) {
                length = std::max(length, std::min(254, cp1252Length(row.get(column))));
            }
            field.length = length;
            field.decimals = 0;
        }
        fields.push_back(std::move(field));
    }
    return fields;
}

std::string formatDbfTextValue(const std::string& value, int length) {
    std::string bytes = toCp1252(value);
    if (static_cast<int>(bytes.size()) > length) {
        bytes.resize(static_cast<size_t>(length));
    }
    if (static_cast<int>(bytes.size()) < length) {
        bytes.append(static_cast<size_t>(length - bytes.size()), ' ');
    }
    return bytes;
}

std::string formatDbfNumberValue(const std::string& value, int length, int decimals) {
    double number = 0.0;
    if (!parseDoubleInvariant(value, number)) {
        return std::string(static_cast<size_t>(length), ' ');
    }
    std::string text = formatDouble(number, decimals);
    if (static_cast<int>(text.size()) > length) {
        text.resize(static_cast<size_t>(length));
    }
    if (static_cast<int>(text.size()) < length) {
        text.insert(text.begin(), static_cast<size_t>(length - text.size()), ' ');
    }
    return text;
}

void writeDbfFile(const std::vector<Row>& rows, const std::vector<DbfField>& fields, const fs::path& path) {
    std::ofstream out(path, std::ios::binary);
    if (!out) {
        throw std::runtime_error("Kan DBF niet schrijven: " + path.string());
    }
    int recordLength = 1;
    for (const auto& field : fields) {
        recordLength += field.length;
    }
    int headerLength = 32 + static_cast<int>(fields.size()) * 32 + 1;

    std::time_t nowTime = std::time(nullptr);
    std::tm now = localTime(nowTime);

    out.put(static_cast<char>(0x03));
    out.put(static_cast<char>(now.tm_year));
    out.put(static_cast<char>(now.tm_mon + 1));
    out.put(static_cast<char>(now.tm_mday));
    writeLEInt32(out, static_cast<int>(rows.size()));
    writeLEInt16(out, headerLength);
    writeLEInt16(out, recordLength);
    for (int i = 0; i < 18; ++i) out.put('\0');
    out.put(static_cast<char>(125));
    out.put('\0');

    for (const auto& field : fields) {
        char name[11] = {};
        std::string cp = toCp1252(field.name);
        std::memcpy(name, cp.data(), std::min<size_t>(10, cp.size()));
        out.write(name, 11);
        out.put(field.type);
        writeLEInt32(out, 0);
        out.put(static_cast<char>(field.length));
        out.put(static_cast<char>(field.decimals));
        for (int i = 0; i < 14; ++i) out.put('\0');
    }
    out.put(static_cast<char>(0x0D));

    for (const auto& row : rows) {
        out.put(' ');
        for (const auto& field : fields) {
            std::string formatted = field.type == 'N'
                ? formatDbfNumberValue(row.get(field.originalName), field.length, field.decimals)
                : formatDbfTextValue(row.get(field.originalName), field.length);
            out.write(formatted.data(), static_cast<std::streamsize>(formatted.size()));
        }
    }
    out.put(static_cast<char>(0x1A));
}

struct ShapefileResult {
    fs::path directory;
    fs::path fieldMapPath;
    bool written = false;
};

ShapefileResult exportPointShapefile(const std::vector<Row>& rows, const fs::path& outputDirectory, const std::string& baseName) {
    std::vector<Row> pointRows;
    for (const auto& row : rows) {
        double x = 0.0;
        double y = 0.0;
        if (parseDoubleInvariant(row.get("FinalRdX"), x) && parseDoubleInvariant(row.get("FinalRdY"), y)) {
            pointRows.push_back(row);
        }
    }
    if (pointRows.empty()) {
        std::cout << "[INFO] Shapefile export overgeslagen: geen rijen met coordinaten.\n";
        return {};
    }

    fs::path shapeDirectory = outputDirectory / (baseName + "_shapefile");
    fs::create_directories(shapeDirectory);
    fs::path shpPath = shapeDirectory / (baseName + ".shp");
    fs::path shxPath = shapeDirectory / (baseName + ".shx");
    fs::path dbfPath = shapeDirectory / (baseName + ".dbf");
    fs::path prjPath = shapeDirectory / (baseName + ".prj");
    fs::path cpgPath = shapeDirectory / (baseName + ".cpg");
    fs::path fieldMapPath = shapeDirectory / (baseName + "_fieldmap.csv");

    double minX = 0.0, minY = 0.0, maxX = 0.0, maxY = 0.0;
    bool first = true;
    for (const auto& row : pointRows) {
        double x = 0.0;
        double y = 0.0;
        parseDoubleInvariant(row.get("FinalRdX"), x);
        parseDoubleInvariant(row.get("FinalRdY"), y);
        if (first) {
            minX = maxX = x;
            minY = maxY = y;
            first = false;
        } else {
            minX = std::min(minX, x);
            maxX = std::max(maxX, x);
            minY = std::min(minY, y);
            maxY = std::max(maxY, y);
        }
    }

    int shpFileLengthWords = 50 + static_cast<int>(pointRows.size()) * 14;
    int shxFileLengthWords = 50 + static_cast<int>(pointRows.size()) * 4;
    std::ofstream shp(shpPath, std::ios::binary);
    std::ofstream shx(shxPath, std::ios::binary);
    if (!shp || !shx) {
        throw std::runtime_error("Kan shapefile niet schrijven.");
    }
    writeShapefileHeader(shp, shpFileLengthWords, 1, minX, minY, maxX, maxY);
    writeShapefileHeader(shx, shxFileLengthWords, 1, minX, minY, maxX, maxY);

    int offsetWords = 50;
    int recordNumber = 1;
    for (const auto& row : pointRows) {
        double x = 0.0;
        double y = 0.0;
        parseDoubleInvariant(row.get("FinalRdX"), x);
        parseDoubleInvariant(row.get("FinalRdY"), y);

        writeBEInt32(shx, offsetWords);
        writeBEInt32(shx, 10);
        writeBEInt32(shp, recordNumber);
        writeBEInt32(shp, 10);
        writeLEInt32(shp, 1);
        writeLEDouble(shp, x);
        writeLEDouble(shp, y);
        offsetWords += 14;
        ++recordNumber;
    }
    shp.close();
    shx.close();

    std::vector<DbfField> fields = buildDbfFields(pointRows);
    writeDbfFile(pointRows, fields, dbfPath);

    std::ostringstream fieldMap;
    fieldMap << csvEscape("Name") << ',' << csvEscape("OriginalName") << ',' << csvEscape("Type") << ','
             << csvEscape("Length") << ',' << csvEscape("Decimals") << "\r\n";
    for (const auto& field : fields) {
        fieldMap << csvEscape(field.name) << ',' << csvEscape(field.originalName) << ',' << csvEscape(std::string(1, field.type))
                 << ',' << csvEscape(std::to_string(field.length)) << ',' << csvEscape(std::to_string(field.decimals)) << "\r\n";
    }
    writeTextUtf8(fieldMapPath, fieldMap.str(), true);
    writeTextUtf8(cpgPath, "1252\r\n", false);
    writeTextUtf8(prjPath, std::string(kRdNewPrj) + "\r\n", false);

    return {shapeDirectory, fieldMapPath, true};
}

std::string timestampText() {
    std::time_t nowTime = std::time(nullptr);
    std::tm now = localTime(nowTime);
    char buffer[32];
    std::strftime(buffer, sizeof(buffer), "%Y%m%d_%H%M%S", &now);
    return buffer;
}

std::string baseNameFromPath(const fs::path& path) {
    return path.stem().string();
}

std::string normalizedArgKey(std::string key) {
    while (!key.empty() && key[0] == '-') {
        key.erase(key.begin());
    }
    key = toLowerAscii(key);
    key = replaceAll(key, "_", "");
    key = replaceAll(key, "-", "");
    return key;
}

bool isSwitchKey(const std::string& key) {
    static const std::unordered_set<std::string> switches = {
        "skipgeocoding", "skipwfs", "help", "h", "?"
    };
    return switches.find(normalizedArgKey(key)) != switches.end();
}

Options parseOptions(int argc, char** argv) {
    Options options;
    for (int i = 1; i < argc; ++i) {
        std::string arg = argv[i];
        std::string key = normalizedArgKey(arg);
        auto nextValue = [&]() -> std::string {
            if (i + 1 >= argc) {
                throw std::runtime_error("Ontbrekende waarde voor argument: " + arg);
            }
            return argv[++i];
        };

        if (key == "help" || key == "h" || key == "?") {
            options.help = true;
        } else if (key == "inputtextpath" || key == "inputtext") {
            options.inputTextPath = nextValue();
        } else if (key == "inputcsvpath" || key == "inputcsv") {
            options.inputCsvPath = nextValue();
        } else if (key == "inputexcelpath" || key == "inputexcel") {
            options.inputExcelPath = nextValue();
        } else if (key == "worksheetname") {
            options.worksheetName = nextValue();
        } else if (key == "outputdirectory" || key == "output") {
            options.outputDirectory = nextValue();
        } else if (key == "wfsurl") {
            options.wfsUrl = nextValue();
        } else if (key == "wfslayername") {
            options.wfsLayerName = nextValue();
        } else if (key == "wfsgeojsonpath") {
            options.wfsGeoJsonPath = nextValue();
        } else if (key == "geocodedelayms") {
            options.geocodeDelayMs = std::stoi(nextValue());
        } else if (key == "geocodeconcurrency") {
            options.geocodeConcurrency = std::stoi(nextValue());
        } else if (key == "snapdistancemeters") {
            options.snapDistanceMeters = std::stoi(nextValue());
        } else if (key == "streetfurnituresnapdistancemeters") {
            options.streetFurnitureSnapDistanceMeters = std::stoi(nextValue());
        } else if (key == "wfsbuffermeters") {
            options.wfsBufferMeters = std::stoi(nextValue());
        } else if (key == "maxrows") {
            options.maxRows = std::stoi(nextValue());
        } else if (key == "skipgeocoding") {
            options.skipGeocoding = true;
        } else if (key == "skipwfs") {
            options.skipWfs = true;
        } else if (!arg.empty() && arg[0] != '-') {
            if (options.inputTextPath.empty() && fs::path(arg).extension() == ".txt") {
                options.inputTextPath = arg;
            } else if (options.inputCsvPath.empty()) {
                options.inputCsvPath = arg;
            } else {
                throw std::runtime_error("Onbekend positioneel argument: " + arg);
            }
        } else if (isSwitchKey(arg)) {
            // Handled above; kept for readability.
        } else {
            throw std::runtime_error("Onbekend argument: " + arg);
        }
    }
    return options;
}

void printUsage() {
    std::cout
        << "Lovion Coordinate Tool C++\n\n"
        << "Gebruik:\n"
        << "  LovionCoordinateToolCpp.exe -InputTextPath export.txt [-OutputDirectory output]\n"
        << "  LovionCoordinateToolCpp.exe -InputCsvPath lovion.csv [-SkipWfs] [-MaxRows 25]\n\n"
        << "Belangrijk:\n"
        << "  Deze native C++ versie leest TXT en CSV. Sla Excel-lijsten eerst op als CSV.\n"
        << "  De bestaande PowerShell-tool blijft beschikbaar voor directe .xlsx-bewerking en de GUI.\n";
}

void runTextMode(const Options& options, const fs::path& outputDirectory) {
    std::vector<Row> rows = readLovionTextRows(options.inputTextPath, options.maxRows);
    std::cout << "[INFO] Records gelezen: " << rows.size() << "\n";

    std::string stamp = timestampText();
    std::string base = baseNameFromPath(options.inputTextPath);
    std::string outBase = base + "_" + stamp;
    fs::path csvPath = outputDirectory / (outBase + ".csv");
    fs::path geoJsonPath = outputDirectory / (outBase + ".geojson");

    exportCsv(rows, csvPath);
    exportGeoJson(rows, geoJsonPath);
    ShapefileResult shape = exportPointShapefile(rows, outputDirectory, outBase);

    int matched = 0;
    for (const auto& row : rows) {
        if (row.get("CoordinateSource") != "none") {
            ++matched;
        }
    }

    std::cout << "\nKlaar.\n";
    std::cout << "Records totaal          : " << rows.size() << "\n";
    std::cout << "Records met coordinaten : " << matched << "\n";
    std::cout << "CSV output              : " << csvPath.string() << "\n";
    std::cout << "GeoJSON output          : " << geoJsonPath.string() << "\n";
    if (shape.written) {
        std::cout << "Shapefile output        : " << shape.directory.string() << "\n";
        std::cout << "Fieldmap output         : " << shape.fieldMapPath.string() << "\n";
    }
}

void runCsvMode(const Options& options, const fs::path& outputDirectory, const std::string& inputPath) {
    std::vector<Row> rows = readCsvRows(inputPath, options.maxRows);
    std::cout << "[INFO] Datarijen gelezen: " << rows.size() << "\n";
    for (auto& row : rows) {
        initializeCoordinateColumns(row);
    }

    if (!options.skipGeocoding) {
        geocodeRowsViaPdok(rows, options);
    } else {
        std::cout << "[INFO] PDOK geocoding overgeslagen.\n";
    }

    std::vector<WfsFeature> wfsPrepared;
    if (!options.skipWfs) {
        try {
            std::vector<Json> rawFeatures;
            if (!options.wfsGeoJsonPath.empty()) {
                std::cout << "[INFO] Laad lokale GeoJSON: " << options.wfsGeoJsonPath << "\n";
                rawFeatures = featuresFromGeoJsonText(readFileUtf8(options.wfsGeoJsonPath));
            } else {
                BoundingBox box = boundingBoxFromPdokRows(rows, options.wfsBufferMeters);
                if (!box.valid) {
                    std::cerr << "[WARN] WFS ophalen overgeslagen: geen RD-coordinaten beschikbaar om een bbox te bouwen.\n";
                } else {
                    rawFeatures = loadWfsFeaturesFromService(options.wfsUrl, options.wfsLayerName, box);
                }
            }
            wfsPrepared = prepareWfsFeatures(rawFeatures);
            std::cout << "[INFO] WFS-features bruikbaar voor matching: " << wfsPrepared.size() << "\n";
        } catch (const std::exception& ex) {
            std::cerr << "[WARN] WFS matching overgeslagen: " << ex.what() << "\n";
        }
    } else {
        std::cout << "[INFO] WFS matching overgeslagen.\n";
    }

    if (!wfsPrepared.empty()) {
        matchRowsToWfs(rows, wfsPrepared, options);
    }

    std::string stamp = timestampText();
    std::string base = baseNameFromPath(inputPath);
    std::string outBase = base + "_" + stamp;
    fs::path csvPath = outputDirectory / (outBase + ".csv");
    fs::path geoJsonPath = outputDirectory / (outBase + ".geojson");

    exportCsv(rows, csvPath);
    exportGeoJson(rows, geoJsonPath);
    ShapefileResult shape = exportPointShapefile(rows, outputDirectory, outBase);

    int matched = 0;
    int wfs = 0;
    int pdok = 0;
    int pdokTolerant = 0;
    for (const auto& row : rows) {
        std::string source = row.get("CoordinateSource");
        if (source != "none") ++matched;
        if (source == "wfs") ++wfs;
        if (source == "pdok") ++pdok;
        if (source == "pdok_tolerant") ++pdokTolerant;
    }

    std::cout << "\nKlaar.\n";
    std::cout << "Rijen totaal            : " << rows.size() << "\n";
    std::cout << "Rijen met coordinaten   : " << matched << "\n";
    std::cout << "Waarvan uit WFS         : " << wfs << "\n";
    std::cout << "Waarvan uit PDOK exact  : " << pdok << "\n";
    std::cout << "Waarvan uit PDOK tol.   : " << pdokTolerant << "\n";
    std::cout << "CSV output              : " << csvPath.string() << "\n";
    std::cout << "GeoJSON output          : " << geoJsonPath.string() << "\n";
    if (shape.written) {
        std::cout << "Shapefile output        : " << shape.directory.string() << "\n";
        std::cout << "Fieldmap output         : " << shape.fieldMapPath.string() << "\n";
    }
}

} // namespace

int main(int argc, char** argv) {
    try {
        Options options = parseOptions(argc, argv);
        if (options.help) {
            printUsage();
            return 0;
        }

        int inputCount = 0;
        if (!options.inputTextPath.empty()) ++inputCount;
        if (!options.inputCsvPath.empty()) ++inputCount;
        if (!options.inputExcelPath.empty()) ++inputCount;
        if (inputCount != 1) {
            printUsage();
            return 2;
        }

        if (!options.inputExcelPath.empty()) {
            fs::path excelPath(options.inputExcelPath);
            if (iequals(excelPath.extension().string(), ".csv")) {
                options.inputCsvPath = options.inputExcelPath;
                options.inputExcelPath.clear();
            } else {
                std::cerr << "Deze C++ versie ondersteunt nog geen directe .xlsx-invoer. "
                          << "Sla het werkblad op als CSV of gebruik de bestaande PowerShell-versie voor Excel/GUI.\n";
                return 2;
            }
        }

        fs::path outputDirectory = options.outputDirectory.empty()
            ? (fs::current_path() / "output_cpp")
            : fs::path(options.outputDirectory);
        fs::create_directories(outputDirectory);

        if (!options.inputTextPath.empty()) {
            runTextMode(options, outputDirectory);
        } else {
            runCsvMode(options, outputDirectory, options.inputCsvPath);
        }
        return 0;
    } catch (const std::exception& ex) {
        std::cerr << "Fout: " << ex.what() << "\n";
        return 1;
    }
}
