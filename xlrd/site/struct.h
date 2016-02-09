#pragma once

#include <string>
#include <vector>
#include <tuple>

namespace xlrd {
namespace site {
namespace structs {  // avoid reserved word

inline
uint8_t
unpack_leB(std::vector<uint8_t> vec, int pos = 0) {
    return vec[pos];
}

inline
uint16_t
unpack_leH(std::vector<uint8_t> vec, int pos = 0) {
    return vec[pos] | (vec[pos+1] << 8);
}

inline
std::tuple<uint16_t, uint16_t> 
unpack_leHH(std::vector<uint8_t> vec, int pos = 0) {
    return std::make_tuple(vec[pos] | (vec[pos+1] << 8),
                           vec[pos+2] | (vec[pos+3] << 8));
}

inline
std::tuple<uint16_t, uint8_t> 
unpack_leHB(std::vector<uint8_t> vec, int pos = 0) {
    return std::make_tuple(vec[pos] | (vec[pos+1] << 8),
                           vec[pos+2]);
}


}
}
}
