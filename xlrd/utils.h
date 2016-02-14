#pragma once

#include <algorithm>
#include <vector>
#include <map>
#include <iostream>

#include "./utils/types.h"
#include "./utils/str.h"

namespace utils {
using any = types::any;

template<class T>
int indexof(std::vector<T> vec, const T& val)
{
    auto it = std::find(vec.begin(), vec.end(), val);
    if (it == vec.end()) {
        return -1;
    }
    return std::distance(vec.begin(), it);
}

template<class T>
auto slice(std::vector<T> vec, int start, int stop, int step=1)
-> std::vector<T>
{
    std::vector<T> dest;
    for (int i=start; i < stop; i += step) {
        dest.push_back(vec[i]);
    }
    return dest;
}

bool equals(std::vector<uint8_t> vec, const std::string& str)
{
    // TODO: compare by simd or uint64_t
    size_t vlen = vec.size();
    size_t slen = str.size();
    if (vlen != slen) return false;
    for (size_t i=0; i < vlen; i++) {
        if (vec[i] != str[i]) return false;
    }
    return true;
}

template<class K>
std::string getelse(const std::map<K, std::string>& dict, K key, const char* default_value)
{
    auto it = dict.find(key);
    if (it == dict.end()) {
        return default_value;
    }
    return it->second;
}

template<class K, class V>
auto getelse(const std::map<K, V>& dict, K key, V default_value)
-> V
{
    auto it = dict.find(key);
    if (it == dict.end()) {
        return default_value;
    }
    return it->second;
}

struct unpack {
public:
    const std::vector<uint8_t>& data_;
    int begin_pos_;
    int end_pos_;
    int pos_;

    unpack(const std::vector<uint8_t>& data, int begin_pos, int end_pos)
    : data_(data), begin_pos_(begin_pos), end_pos_(end_pos), pos_(0)
    {}

    ~unpack() {
        if (begin_pos_+pos_ != end_pos_) throw std::runtime_error("rests");
    }

    template<class T>
    auto as() -> T {
        if (begin_pos_+pos_+sizeof(T) > end_pos_) throw std::runtime_error("over");
        T v = T(*reinterpret_cast<T*>(&data_[begin_pos_+pos_]));
        pos_ += sizeof(T);
        return v;
    }
};

uint8_t as_uint8(std::vector<uint8_t> vec, int pos=0) {
    return vec[pos];
}

uint16_t as_uint16(std::vector<uint8_t> vec, int pos=0) {
    // = unpack("<H", vec[pos:])
    return vec[pos] | (vec[pos+1] << 8);
}

uint16_t as_uint16be(std::vector<uint8_t> vec, int pos=0) {
    // = unpack(">H", vec[pos:])
    return (vec[pos] << 8) | vec[pos+1];
}

int16_t as_int16(std::vector<uint8_t> vec, int pos=0) {
    // = unpack("<h", vec[pos:])
    return vec[pos] | (vec[pos+1] << 8);
}

int16_t as_int16be(std::vector<uint8_t> vec, int pos=0) {
    // = unpack(">h", vec[pos:])
    return (vec[pos] << 8) | vec[pos+1];
}

uint32_t as_uint32(std::vector<uint8_t> vec, int pos=0) {
    return vec[pos] | (vec[pos+1] << 8) | (vec[pos+2] << 16) | (vec[pos+3] << 24);
}

uint32_t as_uint32be(std::vector<uint8_t> vec, int pos=0) {
    return (vec[pos] << 24) | (vec[pos+1] << 16) | (vec[pos+2] << 8) | vec[pos+3];
}

int32_t as_int32(std::vector<uint8_t> vec, int pos=0) {
    return vec[pos] | (vec[pos+1] << 8) | (vec[pos+2] << 16) | (vec[pos+3] << 24);
}

int32_t as_int32be(std::vector<uint8_t> vec, int pos=0) {
    return (vec[pos] << 24) | (vec[pos+1] << 16) | (vec[pos+2] << 8) | vec[pos+3];
}

double as_double(std::vector<uint8_t> vec, int pos=0) {
    return *(double*)&vec[pos];
}

template<class ...A>
void printf(A...a) {
    std::cout << utils::str::format(a...) << std::endl;
}

}
