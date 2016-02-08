#include "xlrd.h"

// clang++ -O2 -std=c++11 test.cpp

int main(int argc, char *argv[])
    auto book = xlrd::open_workbook("test.xls");
    return 0;
}
