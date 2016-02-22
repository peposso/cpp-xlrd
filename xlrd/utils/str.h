//  strutil.h
#pragma once

#include <vector>
#include <map>
#include <memory>
#include <string>
#include <algorithm>

namespace utils {
namespace str {

inline
char lowerchar(char c){
    if('A' <= c && c <= 'Z') return c - ('Z'-'z');
    return c;
} 

inline
std::string lower(const std::string& str){
    std::string lowered = str;
    std::transform(lowered.begin(), lowered.end(), lowered.begin(), lowerchar);
    return lowered;
}

inline
char upperchar(char c){
    if('a' <= c && c <= 'z') return c - ('z'-'Z');
    return c;
}

inline
std::string upper(const std::string& str){
    std::string uppered = str;
    std::transform(uppered.begin(), uppered.end(), uppered.begin(), upperchar);
    return uppered;
}


inline
std::vector<std::string> split(const std::string& haystack, const std::string& needle) {
    std::vector<std::string> splitted;
    size_t pos = 0, next = std::string::npos;
    while (true) {
        next = haystack.find(needle, pos);
        if (next == std::string::npos) break;
        splitted.push_back(haystack.substr(pos, next - pos));
		pos = next + needle.size();
    }
    splitted.push_back(haystack.substr(pos, std::string::npos));
    return splitted;
}

inline
std::vector<std::string> split(const std::string& haystack, char needle) {
    char needle_str[2] = {needle, '\0'};
    return split(haystack, std::string(needle_str));
}


class itersplit_iter {
public:
    inline
    itersplit_iter(const std::string& haystack, const std::string& needle)
    : haystack_(haystack)
    , needle_(needle)
    , pos_(0)
    , next_(std::string::npos)
    {};

    inline
    itersplit_iter(size_t next)
    : haystack_("")
    , needle_("")
    , pos_(0)
    , next_(next)
    {};

    inline
    std::string operator *() {
        next_ = haystack_.find(needle_, pos_);
        return haystack_.substr(pos_, next_ - pos_);
    };

    inline
    void operator ++() {
        if (next_ != std::string::npos) pos_ = next_ + needle_.size();
        next_ = haystack_.find(needle_, pos_);
        if (next_ == std::string::npos) next_ = haystack_.size();
    };

    inline
    bool operator !=(itersplit_iter& it) {
        return next_ != it.next_;
    };

    std::string haystack_;
    std::string needle_;
    size_t pos_;
    size_t next_;

private:
    std::string dirname_;
    std::string name_;
    int d_type_;
};

class itersplit {
public:
    inline
    itersplit(const std::string& haystack, const std::string& needle)
    : haystack_(haystack)
    , needle_(needle)
    {};

    inline
    itersplit_iter begin() { return itersplit_iter(haystack_, needle_); };
    inline
    itersplit_iter end() { return itersplit_iter(haystack_.size()); };

private:
    std::string haystack_;
    std::string needle_;
};

inline
itersplit iterline(const std::string& haystack) {
    return itersplit(haystack, "\n");
}

// make_seq<N> : index_seq<0, 1, ..., N>
template<int... i> struct index_seq{ constexpr index_seq(){}; };
template<int i, int... j> struct make_seq : make_seq<i-1, i, j...>{};
template<int... i> struct make_seq<0, i...> : index_seq<0, i...>{};

// template<class R, class A>
// struct typ{};

// template<class R, class A>
// struct {};

template<class ...A>
inline
std::string format_(const char* fmt, A...a){
    int n = ::snprintf(nullptr, 0, fmt, a...);
    std::string buf(0, n + 1);
    ::snprintf(&buf[0], n+1, fmt, a...);
    return buf;
}


template<class ...A>
std::string repr(A...);

template<> std::string repr(const std::string& a) { return format_("\"%s\"", a.c_str()); };
template<> std::string repr(char* a) { return format_("\"%s\"", a); };

template<> std::string repr(int8_t a) { return format_("%d", a); };
template<> std::string repr(int16_t a) { return format_("%d", a); };
template<> std::string repr(int32_t a) { return format_("%d", a); };
template<> std::string repr(int64_t a) { return format_("%ld", a); };

template<> std::string repr(uint8_t a) { return format_("%u", a); }
template<> std::string repr(uint16_t a) { return format_("%u", a); }
template<> std::string repr(uint32_t a) { return format_("%u", a); }
template<> std::string repr(uint64_t a) { return format_("%lu", a); }

template<> std::string repr(double a) { return format_("%f", a); }
template<> std::string repr(float a) { return format_("%f", a); }

template<class F, class ...Rest>
std::string repr(F f, Rest...r) {
    return format_("%s, %s", repr(f).c_str(), repr(r).c_str()...);
}

template<class V>
std::string repr(const std::vector<V>& vec) {
    std::string buf = "vector{";
    int len = (int)vec.size();
    for (int i=0; i < len; i++) {
        buf.append(repr(vec[i]));
        if (i < len-1) buf.append(", ");
    }
    buf.push_back('}');
    return buf;
}

template<class K, class V>
std::string repr(const std::map<K, V>& m) {
    std::string buf = "map{";
    size_t len = m.size();
    size_t i = 0;
    for (auto kv: m) {
        buf.append(repr(kv->first));
        buf.append(": ");
        buf.append(repr(kv->second));
        if (i++ < len-1) buf.append(", ");
    }
    buf.push_back('}');
    return buf;
}

template<class ...T, int... I>
std::string repr_tuple_impl(std::tuple<T...> t, index_seq<I...>) {
    std::string buf = "tuple(";
    buf.append(repr(std::get<I>(t)...));
    buf.push_back(')');
    return buf;
}

template<class ...T>
std::string repr(std::tuple<T...> t) {
    return repr_tuple_impl(t, make_seq<sizeof...(T)-1>{});
}

template<class A, class R>
auto tocharptr(A a) -> R;

template<>
auto tocharptr(const std::string& a)
-> const char* {
    return a.c_str();
};

template<class V>
auto tocharptr(const std::vector<V>& a)
-> const char* {
    return repr(a).c_str();
};

template<class A>
auto tocharptr(A a) -> A {
    return a;
};

template<class ...A>
inline
std::string format(const char* fmt, A...a){
    return format(fmt, tocharptr(a)...);
}

inline
std::string replace(std::string src, const std::string& from, const std::string& to) {
    std::string dest = src;
    std::string::size_type pos = 0;
    while(pos = dest.find(from, pos), pos != std::string::npos) {
        dest.replace(pos, from.length(), to);
        pos += to.length();
    }
    return dest;
}

std::string utf16to8(const std::vector<uint8_t>& u16buf) {
    std::string u8buf = "";
    for (uint32_t i=0; i < u16buf.size(); i+=2) {
        int uc = u16buf[i] | (u16buf[i+1] << 8);
        if (uc < 0x7f) {
            // ascii
            u8buf.push_back(uc);
        } else if (uc < 0x7FF) {
            // 2bytes
            uint8_t b1 = 0xC2 | (0b00011111 & (uc>>6));
            uint8_t b2 = 0x80 | (0b00111111 & uc);
            u8buf.push_back(b1);
            u8buf.push_back(b2);
        } else if (uc < 0xFFFF) {
            // 3bytes
            uint8_t b1 = 0xE0 | (0b00001111 & (uc>>12));
            uint8_t b2 = 0x80 | (0b00111111 & (uc>>6));
            uint8_t b3 = 0x80 | (0b00111111 & uc);
            u8buf.push_back(b1);
            u8buf.push_back(b2);
            u8buf.push_back(b3);
        }
    }
    return u8buf;
}

std::string unicode(std::vector<uint8_t> src, std::string encoding)
{
    return std::string((char*)&src[0], src.size());
}

std::string ltrim(const std::string& src)
{
    size_t pos = 0;
    for (int i=0; i < (int)src.size(); i++) {
        char c = src[i];
        if (c!=' '&&c!='\t'&&c!='\r'&&c!='\n') {
            pos = i;
            break;
        }
    }
    return src.substr(pos);
}

std::string rtrim(const std::string& src)
{
    size_t pos = src.size()-1;
    for (int i=src.size()-1; i >= 0; i--) {
        char c = src[i];
        if (c!=' '&&c!='\t'&&c!='\r'&&c!='\n') {
            pos = i;
            break;
        }
    }
    return src.substr(0, pos+1);
}

std::string trim(const std::string& src)
{
    return ltrim(rtrim(src));
}

}
}
