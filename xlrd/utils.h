#pragma once

#include <algorithm>
#include <vector>

namespace xlrd {
namespace utils {

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

std::string slice_as_str(std::vector<uint8_t> vec, int start, int stop, int step=1)
{
    std::string dest;
    for (int i=start; i < stop; i += step) {
        dest.push_back(vec[i]);
    }
    return dest;
}

uint8_t as_uint8(std::vector<uint8_t> vec, int pos=0) {
    return vec[pos];
}

uint16_t as_uint16le(std::vector<uint8_t> vec, int pos=0) {
    // = unpack("<H", vec[pos:])
    return vec[pos] | (vec[pos+1] << 8);
}

uint16_t as_uint16be(std::vector<uint8_t> vec, int pos=0) {
    // = unpack(">H", vec[pos:])
    return (vec[pos] << 8) | vec[pos+1];
}

int16_t as_int16le(std::vector<uint8_t> vec, int pos=0) {
    // = unpack("<h", vec[pos:])
    return vec[pos] | (vec[pos+1] << 8);
}

int16_t as_int16be(std::vector<uint8_t> vec, int pos=0) {
    // = unpack(">h", vec[pos:])
    return (vec[pos] << 8) | vec[pos+1];
}

uint32_t as_uint32le(std::vector<uint8_t> vec, int pos=0) {
    return vec[pos] | (vec[pos+1] << 8) | (vec[pos+2] << 16) | (vec[pos+3] << 24);
}

uint32_t as_uint32be(std::vector<uint8_t> vec, int pos=0) {
    return (vec[pos] << 24) | (vec[pos+1] << 16) | (vec[pos+2] << 8) | vec[pos+3];
}

int32_t as_int32le(std::vector<uint8_t> vec, int pos=0) {
    return vec[pos] | (vec[pos+1] << 8) | (vec[pos+2] << 16) | (vec[pos+3] << 24);
}

int32_t as_int32be(std::vector<uint8_t> vec, int pos=0) {
    return (vec[pos] << 24) | (vec[pos+1] << 16) | (vec[pos+2] << 8) | vec[pos+3];
}

std::string utf16to8(std::vector<uint8_t> u16buf) {
    std::string u8buf = "";
    for (int i=0; i < u16buf.size(); i+=2) {
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

}
}
