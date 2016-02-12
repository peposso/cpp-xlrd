#pragma once

#include <algorithm>
#include <vector>
#include <map>

#include "./utils/types.h"
#include "./utils/str.h"

namespace utils {
using any = types::any;

template<class T>
size_t indexof(std::vector<T> vec, const T& val)
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

}
