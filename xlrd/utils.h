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




}
}
