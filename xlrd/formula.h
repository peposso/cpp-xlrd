#pragma once

// -*- coding: cp1252 -*-

////
// Module for parsing/evaluating Microsoft Excel formulas.
//
// <p>Copyright � 2005-2012 Stephen John Machin, Lingfo Pty Ltd</p>
// <p>This module is part of the xlrd package, which is released under
// a BSD-style licence.</p>
////

// No part of the content of this file was derived from the works of David Giffin.

#include <set>

#include "./biffh.h"  //  unpack_unicode_update_pos, unpack_string_update_pos, \
                    //  XLRDError, hex_char_dump, error_text_from_code, BaseObject

#include "./utils.h"

namespace xlrd {
namespace formula {

namespace strutil = utils::str;

auto& unpack_unicode_update_pos = biffh::unpack_unicode_update_pos;
auto& unpack_string_update_pos = biffh::unpack_string_update_pos;
using XLRDError = biffh::XLRDError;
auto& error_text_from_code = biffh::error_text_from_code;

/*
from .biffh import unpack_unicode_update_pos, unpack_string_update_pos, \
    XLRDError, hex_char_dump, error_text_from_code, BaseObject

__all__ = [
    'oBOOL', 'oERR', 'oNUM', 'oREF', 'oREL', 'oSTRG', 'oUNK',
    'decompile_formula',
    'dump_formula',
    'evaluate_name_formula',
    'okind_dict',
    'rangename3d', 'rangename3drel', 'cellname', 'cellnameabs', 'colname',
    'FMLA_TYPE_CELL',
    'FMLA_TYPE_SHARED',
    'FMLA_TYPE_ARRAY',
    'FMLA_TYPE_COND_FMT',
    'FMLA_TYPE_DATA_VAL',
    'FMLA_TYPE_NAME',
    ]
*/

static int FMLA_TYPE_CELL = 1;
static int FMLA_TYPE_SHARED = 2;
static int FMLA_TYPE_ARRAY = 4;
static int FMLA_TYPE_COND_FMT = 8;
static int FMLA_TYPE_DATA_VAL = 16;
static int FMLA_TYPE_NAME = 32;
static int ALL_FMLA_TYPES = 63;

static std::map<int, std::string>
FMLA_TYPEDESCR_MAP = {
    {1 , "CELL"},
    {2 , "SHARED"},
    {4 , "ARRAY"},
    {8 , "COND-FMT"},
    {16, "DATA-VAL"},
    {32, "NAME"},
};

static std::map<int, int>
_TOKEN_NOT_ALLOWED_DICT = {
    {0x01,   ALL_FMLA_TYPES - FMLA_TYPE_CELL}, // tExp
    {0x02,   ALL_FMLA_TYPES - FMLA_TYPE_CELL}, // tTbl
    {0x0F,   FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL}, // tIsect
    {0x10,   FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL}, // tUnion/List
    {0x11,   FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL}, // tRange
    {0x20,   FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL}, // tArray
    {0x23,   FMLA_TYPE_SHARED}, // tName
    {0x39,   FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL}, // tNameX
    {0x3A,   FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL}, // tRef3d
    {0x3B,   FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL}, // tArea3d
    {0x2C,   FMLA_TYPE_CELL + FMLA_TYPE_ARRAY}, // tRefN
    {0x2D,   FMLA_TYPE_CELL + FMLA_TYPE_ARRAY}, // tAreaN
    // plus weird stuff like tMem*
};

inline
int _TOKEN_NOT_ALLOWED(int key, int alt) {
    auto it = _TOKEN_NOT_ALLOWED_DICT.find(key);
    if (it == _TOKEN_NOT_ALLOWED_DICT.end()) {
        return it->second;
    }
    return alt;
}


static int oBOOL = 3;
static int oERR =  4;
static int oMSNG = 5; // tMissArg
static int oNUM =  2;
static int oREF = -1;
static int oREL = -2;
static int oSTRG = 1;
static int oUNK =  0;

static std::map<int, std::string>
okind_dict = {
    {-2, "oREL"},
    {-1, "oREF"},
    {0 , "oUNK"},
    {1 , "oSTRG"},
    {2 , "oNUM"},
    {3 , "oBOOL"},
    {4 , "oERR"},
    {5 , "oMSNG"},
};

static char
listsep = ','; //////// probably should depend on locale


// sztabN[opcode] -> the number of bytes to consume.
// -1 means variable
// -2 means this opcode not implemented in this version.
// Which N to use? Depends on biff_version; see szdict.
static std::vector<int>
sztab0 = {-2, 4, 4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -2, -1, 8, 4, 2, 2, 3, 9, 8, 2, 3, 8, 4, 7, 5, 5, 5, 2, 4, 7, 4, 7, 2, 2, -2, -2, -2, -2, -2, -2, -2, -2, 3, -2, -2, -2, -2, -2, -2, -2};
static std::vector<int>
sztab1 = {-2, 5, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -2, -1, 11, 5, 2, 2, 3, 9, 9, 2, 3, 11, 4, 7, 7, 7, 7, 3, 4, 7, 4, 7, 3, 3, -2, -2, -2, -2, -2, -2, -2, -2, 3, -2, -2, -2, -2, -2, -2, -2};
static std::vector<int>
sztab2 = {-2, 5, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -2, -1, 11, 5, 2, 2, 3, 9, 9, 3, 4, 11, 4, 7, 7, 7, 7, 3, 4, 7, 4, 7, 3, 3, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2};
static std::vector<int>
sztab3 = {-2, 5, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -2, -1, -2, -2, 2, 2, 3, 9, 9, 3, 4, 15, 4, 7, 7, 7, 7, 3, 4, 7, 4, 7, 3, 3, -2, -2, -2, -2, -2, -2, -2, -2, -2, 25, 18, 21, 18, 21, -2, -2};
static std::vector<int>
sztab4 = {-2, 5, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -1, -1, -2, -2, 2, 2, 3, 9, 9, 3, 4, 5, 5, 9, 7, 7, 7, 3, 5, 9, 5, 9, 3, 3, -2, -2, -2, -2, -2, -2, -2, -2, -2, 7, 7, 11, 7, 11, -2, -2};

static std::map<int, std::vector<int>&>
szdict = {
    {20, sztab0},
    {21, sztab0},
    {30, sztab1},
    {40, sztab2},
    {45, sztab2},
    {50, sztab3},
    {70, sztab3},
    {80, sztab4},
};

// For debugging purposes ... the name for each opcode
// (without the prefix "t" used on OOo docs)
static std::vector<std::string>
onames = {"Unk00", "Exp", "Tbl", "Add", "Sub", "Mul", "Div", "Power", "Concat", "LT", "LE", "EQ", "GE", "GT", "NE", "Isect", "List", "Range", "Uplus", "Uminus", "Percent", "Paren", "MissArg", "Str", "Extended", "Attr", "Sheet", "EndSheet", "Err", "Bool", "Int", "Num", "Array", "Func", "FuncVar", "Name", "Ref", "Area", "MemArea", "MemErr", "MemNoMem", "MemFunc", "RefErr", "AreaErr", "RefN", "AreaN", "MemAreaN", "MemNoMemN", "", "", "", "", "", "", "", "", "FuncCE", "NameX", "Ref3d", "Area3d", "RefErr3d", "AreaErr3d", "", ""};

static std::map<int, std::tuple<std::string, int, int, int, int, std::string, std::string>>
#define T std::make_tuple
func_defs = {
    // index: (name, min//args, max//args, flags, //known_args, return_type, kargs)
    {0  , T("COUNT",            0, 30, 0x04,  1, "V", "R")},
    {1  , T("IF",               2,  3, 0x04,  3, "V", "VRR")},
    {2  , T("ISNA",             1,  1, 0x02,  1, "V", "V")},
    {3  , T("ISERROR",          1,  1, 0x02,  1, "V", "V")},
    {4  , T("SUM",              0, 30, 0x04,  1, "V", "R")},
    {5  , T("AVERAGE",          1, 30, 0x04,  1, "V", "R")},
    {6  , T("MIN",              1, 30, 0x04,  1, "V", "R")},
    {7  , T("MAX",              1, 30, 0x04,  1, "V", "R")},
    {8  , T("ROW",              0,  1, 0x04,  1, "V", "R")},
    {9  , T("COLUMN",           0,  1, 0x04,  1, "V", "R")},
    {10 , T("NA",               0,  0, 0x02,  0, "V", "")},
    {11 , T("NPV",              2, 30, 0x04,  2, "V", "VR")},
    {12 , T("STDEV",            1, 30, 0x04,  1, "V", "R")},
    {13 , T("DOLLAR",           1,  2, 0x04,  1, "V", "V")},
    {14 , T("FIXED",            2,  3, 0x04,  3, "V", "VVV")},
    {15 , T("SIN",              1,  1, 0x02,  1, "V", "V")},
    {16 , T("COS",              1,  1, 0x02,  1, "V", "V")},
    {17 , T("TAN",              1,  1, 0x02,  1, "V", "V")},
    {18 , T("ATAN",             1,  1, 0x02,  1, "V", "V")},
    {19 , T("PI",               0,  0, 0x02,  0, "V", "")},
    {20 , T("SQRT",             1,  1, 0x02,  1, "V", "V")},
    {21 , T("EXP",              1,  1, 0x02,  1, "V", "V")},
    {22 , T("LN",               1,  1, 0x02,  1, "V", "V")},
    {23 , T("LOG10",            1,  1, 0x02,  1, "V", "V")},
    {24 , T("ABS",              1,  1, 0x02,  1, "V", "V")},
    {25 , T("INT",              1,  1, 0x02,  1, "V", "V")},
    {26 , T("SIGN",             1,  1, 0x02,  1, "V", "V")},
    {27 , T("ROUND",            2,  2, 0x02,  2, "V", "VV")},
    {28 , T("LOOKUP",           2,  3, 0x04,  2, "V", "VR")},
    {29 , T("INDEX",            2,  4, 0x0c,  4, "R", "RVVV")},
    {30 , T("REPT",             2,  2, 0x02,  2, "V", "VV")},
    {31 , T("MID",              3,  3, 0x02,  3, "V", "VVV")},
    {32 , T("LEN",              1,  1, 0x02,  1, "V", "V")},
    {33 , T("VALUE",            1,  1, 0x02,  1, "V", "V")},
    {34 , T("TRUE",             0,  0, 0x02,  0, "V", "")},
    {35 , T("FALSE",            0,  0, 0x02,  0, "V", "")},
    {36 , T("AND",              1, 30, 0x04,  1, "V", "R")},
    {37 , T("OR",               1, 30, 0x04,  1, "V", "R")},
    {38 , T("NOT",              1,  1, 0x02,  1, "V", "V")},
    {39 , T("MOD",              2,  2, 0x02,  2, "V", "VV")},
    {40 , T("DCOUNT",           3,  3, 0x02,  3, "V", "RRR")},
    {41 , T("DSUM",             3,  3, 0x02,  3, "V", "RRR")},
    {42 , T("DAVERAGE",         3,  3, 0x02,  3, "V", "RRR")},
    {43 , T("DMIN",             3,  3, 0x02,  3, "V", "RRR")},
    {44 , T("DMAX",             3,  3, 0x02,  3, "V", "RRR")},
    {45 , T("DSTDEV",           3,  3, 0x02,  3, "V", "RRR")},
    {46 , T("VAR",              1, 30, 0x04,  1, "V", "R")},
    {47 , T("DVAR",             3,  3, 0x02,  3, "V", "RRR")},
    {48 , T("TEXT",             2,  2, 0x02,  2, "V", "VV")},
    {49 , T("LINEST",           1,  4, 0x04,  4, "A", "RRVV")},
    {50 , T("TREND",            1,  4, 0x04,  4, "A", "RRRV")},
    {51 , T("LOGEST",           1,  4, 0x04,  4, "A", "RRVV")},
    {52 , T("GROWTH",           1,  4, 0x04,  4, "A", "RRRV")},
    {56 , T("PV",               3,  5, 0x04,  5, "V", "VVVVV")},
    {57 , T("FV",               3,  5, 0x04,  5, "V", "VVVVV")},
    {58 , T("NPER",             3,  5, 0x04,  5, "V", "VVVVV")},
    {59 , T("PMT",              3,  5, 0x04,  5, "V", "VVVVV")},
    {60 , T("RATE",             3,  6, 0x04,  6, "V", "VVVVVV")},
    {61 , T("MIRR",             3,  3, 0x02,  3, "V", "RVV")},
    {62 , T("IRR",              1,  2, 0x04,  2, "V", "RV")},
    {63 , T("RAND",             0,  0, 0x0a,  0, "V", "")},
    {64 , T("MATCH",            2,  3, 0x04,  3, "V", "VRR")},
    {65 , T("DATE",             3,  3, 0x02,  3, "V", "VVV")},
    {66 , T("TIME",             3,  3, 0x02,  3, "V", "VVV")},
    {67 , T("DAY",              1,  1, 0x02,  1, "V", "V")},
    {68 , T("MONTH",            1,  1, 0x02,  1, "V", "V")},
    {69 , T("YEAR",             1,  1, 0x02,  1, "V", "V")},
    {70 , T("WEEKDAY",          1,  2, 0x04,  2, "V", "VV")},
    {71 , T("HOUR",             1,  1, 0x02,  1, "V", "V")},
    {72 , T("MINUTE",           1,  1, 0x02,  1, "V", "V")},
    {73 , T("SECOND",           1,  1, 0x02,  1, "V", "V")},
    {74 , T("NOW",              0,  0, 0x0a,  0, "V", "")},
    {75 , T("AREAS",            1,  1, 0x02,  1, "V", "R")},
    {76 , T("ROWS",             1,  1, 0x02,  1, "V", "R")},
    {77 , T("COLUMNS",          1,  1, 0x02,  1, "V", "R")},
    {78 , T("OFFSET",           3,  5, 0x04,  5, "R", "RVVVV")},
    {82 , T("SEARCH",           2,  3, 0x04,  3, "V", "VVV")},
    {83 , T("TRANSPOSE",        1,  1, 0x02,  1, "A", "A")},
    {86 , T("TYPE",             1,  1, 0x02,  1, "V", "V")},
    {92 , T("SERIESSUM",        4,  4, 0x02,  4, "V", "VVVA")},
    {97 , T("ATAN2",            2,  2, 0x02,  2, "V", "VV")},
    {98 , T("ASIN",             1,  1, 0x02,  1, "V", "V")},
    {99 , T("ACOS",             1,  1, 0x02,  1, "V", "V")},
    {100, T("CHOOSE",           2, 30, 0x04,  2, "V", "VR")},
    {101, T("HLOOKUP",          3,  4, 0x04,  4, "V", "VRRV")},
    {102, T("VLOOKUP",          3,  4, 0x04,  4, "V", "VRRV")},
    {105, T("ISREF",            1,  1, 0x02,  1, "V", "R")},
    {109, T("LOG",              1,  2, 0x04,  2, "V", "VV")},
    {111, T("CHAR",             1,  1, 0x02,  1, "V", "V")},
    {112, T("LOWER",            1,  1, 0x02,  1, "V", "V")},
    {113, T("UPPER",            1,  1, 0x02,  1, "V", "V")},
    {114, T("PROPER",           1,  1, 0x02,  1, "V", "V")},
    {115, T("LEFT",             1,  2, 0x04,  2, "V", "VV")},
    {116, T("RIGHT",            1,  2, 0x04,  2, "V", "VV")},
    {117, T("EXACT",            2,  2, 0x02,  2, "V", "VV")},
    {118, T("TRIM",             1,  1, 0x02,  1, "V", "V")},
    {119, T("REPLACE",          4,  4, 0x02,  4, "V", "VVVV")},
    {120, T("SUBSTITUTE",       3,  4, 0x04,  4, "V", "VVVV")},
    {121, T("CODE",             1,  1, 0x02,  1, "V", "V")},
    {124, T("FIND",             2,  3, 0x04,  3, "V", "VVV")},
    {125, T("CELL",             1,  2, 0x0c,  2, "V", "VR")},
    {126, T("ISERR",            1,  1, 0x02,  1, "V", "V")},
    {127, T("ISTEXT",           1,  1, 0x02,  1, "V", "V")},
    {128, T("ISNUMBER",         1,  1, 0x02,  1, "V", "V")},
    {129, T("ISBLANK",          1,  1, 0x02,  1, "V", "V")},
    {130, T("T",                1,  1, 0x02,  1, "V", "R")},
    {131, T("N",                1,  1, 0x02,  1, "V", "R")},
    {140, T("DATEVALUE",        1,  1, 0x02,  1, "V", "V")},
    {141, T("TIMEVALUE",        1,  1, 0x02,  1, "V", "V")},
    {142, T("SLN",              3,  3, 0x02,  3, "V", "VVV")},
    {143, T("SYD",              4,  4, 0x02,  4, "V", "VVVV")},
    {144, T("DDB",              4,  5, 0x04,  5, "V", "VVVVV")},
    {148, T("INDIRECT",         1,  2, 0x0c,  2, "R", "VV")},
    {162, T("CLEAN",            1,  1, 0x02,  1, "V", "V")},
    {163, T("MDETERM",          1,  1, 0x02,  1, "V", "A")},
    {164, T("MINVERSE",         1,  1, 0x02,  1, "A", "A")},
    {165, T("MMULT",            2,  2, 0x02,  2, "A", "AA")},
    {167, T("IPMT",             4,  6, 0x04,  6, "V", "VVVVVV")},
    {168, T("PPMT",             4,  6, 0x04,  6, "V", "VVVVVV")},
    {169, T("COUNTA",           0, 30, 0x04,  1, "V", "R")},
    {183, T("PRODUCT",          0, 30, 0x04,  1, "V", "R")},
    {184, T("FACT",             1,  1, 0x02,  1, "V", "V")},
    {189, T("DPRODUCT",         3,  3, 0x02,  3, "V", "RRR")},
    {190, T("ISNONTEXT",        1,  1, 0x02,  1, "V", "V")},
    {193, T("STDEVP",           1, 30, 0x04,  1, "V", "R")},
    {194, T("VARP",             1, 30, 0x04,  1, "V", "R")},
    {195, T("DSTDEVP",          3,  3, 0x02,  3, "V", "RRR")},
    {196, T("DVARP",            3,  3, 0x02,  3, "V", "RRR")},
    {197, T("TRUNC",            1,  2, 0x04,  2, "V", "VV")},
    {198, T("ISLOGICAL",        1,  1, 0x02,  1, "V", "V")},
    {199, T("DCOUNTA",          3,  3, 0x02,  3, "V", "RRR")},
    {204, T("USDOLLAR",         1,  2, 0x04,  2, "V", "VV")},
    {205, T("FINDB",            2,  3, 0x04,  3, "V", "VVV")},
    {206, T("SEARCHB",          2,  3, 0x04,  3, "V", "VVV")},
    {207, T("REPLACEB",         4,  4, 0x02,  4, "V", "VVVV")},
    {208, T("LEFTB",            1,  2, 0x04,  2, "V", "VV")},
    {209, T("RIGHTB",           1,  2, 0x04,  2, "V", "VV")},
    {210, T("MIDB",             3,  3, 0x02,  3, "V", "VVV")},
    {211, T("LENB",             1,  1, 0x02,  1, "V", "V")},
    {212, T("ROUNDUP",          2,  2, 0x02,  2, "V", "VV")},
    {213, T("ROUNDDOWN",        2,  2, 0x02,  2, "V", "VV")},
    {214, T("ASC",              1,  1, 0x02,  1, "V", "V")},
    {215, T("DBCS",             1,  1, 0x02,  1, "V", "V")},
    {216, T("RANK",             2,  3, 0x04,  3, "V", "VRV")},
    {219, T("ADDRESS",          2,  5, 0x04,  5, "V", "VVVVV")},
    {220, T("DAYS360",          2,  3, 0x04,  3, "V", "VVV")},
    {221, T("TODAY",            0,  0, 0x0a,  0, "V", "")},
    {222, T("VDB",              5,  7, 0x04,  7, "V", "VVVVVVV")},
    {227, T("MEDIAN",           1, 30, 0x04,  1, "V", "R")},
    {228, T("SUMPRODUCT",       1, 30, 0x04,  1, "V", "A")},
    {229, T("SINH",             1,  1, 0x02,  1, "V", "V")},
    {230, T("COSH",             1,  1, 0x02,  1, "V", "V")},
    {231, T("TANH",             1,  1, 0x02,  1, "V", "V")},
    {232, T("ASINH",            1,  1, 0x02,  1, "V", "V")},
    {233, T("ACOSH",            1,  1, 0x02,  1, "V", "V")},
    {234, T("ATANH",            1,  1, 0x02,  1, "V", "V")},
    {235, T("DGET",             3,  3, 0x02,  3, "V", "RRR")},
    {244, T("INFO",             1,  1, 0x02,  1, "V", "V")},
    {247, T("DB",               4,  5, 0x04,  5, "V", "VVVVV")},
    {252, T("FREQUENCY",        2,  2, 0x02,  2, "A", "RR")},
    {261, T("ERROR.TYPE",       1,  1, 0x02,  1, "V", "V")},
    {269, T("AVEDEV",           1, 30, 0x04,  1, "V", "R")},
    {270, T("BETADIST",         3,  5, 0x04,  1, "V", "V")},
    {271, T("GAMMALN",          1,  1, 0x02,  1, "V", "V")},
    {272, T("BETAINV",          3,  5, 0x04,  1, "V", "V")},
    {273, T("BINOMDIST",        4,  4, 0x02,  4, "V", "VVVV")},
    {274, T("CHIDIST",          2,  2, 0x02,  2, "V", "VV")},
    {275, T("CHIINV",           2,  2, 0x02,  2, "V", "VV")},
    {276, T("COMBIN",           2,  2, 0x02,  2, "V", "VV")},
    {277, T("CONFIDENCE",       3,  3, 0x02,  3, "V", "VVV")},
    {278, T("CRITBINOM",        3,  3, 0x02,  3, "V", "VVV")},
    {279, T("EVEN",             1,  1, 0x02,  1, "V", "V")},
    {280, T("EXPONDIST",        3,  3, 0x02,  3, "V", "VVV")},
    {281, T("FDIST",            3,  3, 0x02,  3, "V", "VVV")},
    {282, T("FINV",             3,  3, 0x02,  3, "V", "VVV")},
    {283, T("FISHER",           1,  1, 0x02,  1, "V", "V")},
    {284, T("FISHERINV",        1,  1, 0x02,  1, "V", "V")},
    {285, T("FLOOR",            2,  2, 0x02,  2, "V", "VV")},
    {286, T("GAMMADIST",        4,  4, 0x02,  4, "V", "VVVV")},
    {287, T("GAMMAINV",         3,  3, 0x02,  3, "V", "VVV")},
    {288, T("CEILING",          2,  2, 0x02,  2, "V", "VV")},
    {289, T("HYPGEOMDIST",      4,  4, 0x02,  4, "V", "VVVV")},
    {290, T("LOGNORMDIST",      3,  3, 0x02,  3, "V", "VVV")},
    {291, T("LOGINV",           3,  3, 0x02,  3, "V", "VVV")},
    {292, T("NEGBINOMDIST",     3,  3, 0x02,  3, "V", "VVV")},
    {293, T("NORMDIST",         4,  4, 0x02,  4, "V", "VVVV")},
    {294, T("NORMSDIST",        1,  1, 0x02,  1, "V", "V")},
    {295, T("NORMINV",          3,  3, 0x02,  3, "V", "VVV")},
    {296, T("NORMSINV",         1,  1, 0x02,  1, "V", "V")},
    {297, T("STANDARDIZE",      3,  3, 0x02,  3, "V", "VVV")},
    {298, T("ODD",              1,  1, 0x02,  1, "V", "V")},
    {299, T("PERMUT",           2,  2, 0x02,  2, "V", "VV")},
    {300, T("POISSON",          3,  3, 0x02,  3, "V", "VVV")},
    {301, T("TDIST",            3,  3, 0x02,  3, "V", "VVV")},
    {302, T("WEIBULL",          4,  4, 0x02,  4, "V", "VVVV")},
    {303, T("SUMXMY2",          2,  2, 0x02,  2, "V", "AA")},
    {304, T("SUMX2MY2",         2,  2, 0x02,  2, "V", "AA")},
    {305, T("SUMX2PY2",         2,  2, 0x02,  2, "V", "AA")},
    {306, T("CHITEST",          2,  2, 0x02,  2, "V", "AA")},
    {307, T("CORREL",           2,  2, 0x02,  2, "V", "AA")},
    {308, T("COVAR",            2,  2, 0x02,  2, "V", "AA")},
    {309, T("FORECAST",         3,  3, 0x02,  3, "V", "VAA")},
    {310, T("FTEST",            2,  2, 0x02,  2, "V", "AA")},
    {311, T("INTERCEPT",        2,  2, 0x02,  2, "V", "AA")},
    {312, T("PEARSON",          2,  2, 0x02,  2, "V", "AA")},
    {313, T("RSQ",              2,  2, 0x02,  2, "V", "AA")},
    {314, T("STEYX",            2,  2, 0x02,  2, "V", "AA")},
    {315, T("SLOPE",            2,  2, 0x02,  2, "V", "AA")},
    {316, T("TTEST",            4,  4, 0x02,  4, "V", "AAVV")},
    {317, T("PROB",             3,  4, 0x04,  3, "V", "AAV")},
    {318, T("DEVSQ",            1, 30, 0x04,  1, "V", "R")},
    {319, T("GEOMEAN",          1, 30, 0x04,  1, "V", "R")},
    {320, T("HARMEAN",          1, 30, 0x04,  1, "V", "R")},
    {321, T("SUMSQ",            0, 30, 0x04,  1, "V", "R")},
    {322, T("KURT",             1, 30, 0x04,  1, "V", "R")},
    {323, T("SKEW",             1, 30, 0x04,  1, "V", "R")},
    {324, T("ZTEST",            2,  3, 0x04,  2, "V", "RV")},
    {325, T("LARGE",            2,  2, 0x02,  2, "V", "RV")},
    {326, T("SMALL",            2,  2, 0x02,  2, "V", "RV")},
    {327, T("QUARTILE",         2,  2, 0x02,  2, "V", "RV")},
    {328, T("PERCENTILE",       2,  2, 0x02,  2, "V", "RV")},
    {329, T("PERCENTRANK",      2,  3, 0x04,  2, "V", "RV")},
    {330, T("MODE",             1, 30, 0x04,  1, "V", "A")},
    {331, T("TRIMMEAN",         2,  2, 0x02,  2, "V", "RV")},
    {332, T("TINV",             2,  2, 0x02,  2, "V", "VV")},
    {336, T("CONCATENATE",      0, 30, 0x04,  1, "V", "V")},
    {337, T("POWER",            2,  2, 0x02,  2, "V", "VV")},
    {342, T("RADIANS",          1,  1, 0x02,  1, "V", "V")},
    {343, T("DEGREES",          1,  1, 0x02,  1, "V", "V")},
    {344, T("SUBTOTAL",         2, 30, 0x04,  2, "V", "VR")},
    {345, T("SUMIF",            2,  3, 0x04,  3, "V", "RVR")},
    {346, T("COUNTIF",          2,  2, 0x02,  2, "V", "RV")},
    {347, T("COUNTBLANK",       1,  1, 0x02,  1, "V", "R")},
    {350, T("ISPMT",            4,  4, 0x02,  4, "V", "VVVV")},
    {351, T("DATEDIF",          3,  3, 0x02,  3, "V", "VVV")},
    {352, T("DATESTRING",       1,  1, 0x02,  1, "V", "V")},
    {353, T("NUMBERSTRING",     2,  2, 0x02,  2, "V", "VV")},
    {354, T("ROMAN",            1,  2, 0x04,  2, "V", "VV")},
    {358, T("GETPIVOTDATA",     2,  2, 0x02,  2, "V", "RV")},
    {359, T("HYPERLINK",        1,  2, 0x04,  2, "V", "VV")},
    {360, T("PHONETIC",         1,  1, 0x02,  1, "V", "V")},
    {361, T("AVERAGEA",         1, 30, 0x04,  1, "V", "R")},
    {362, T("MAXA",             1, 30, 0x04,  1, "V", "R")},
    {363, T("MINA",             1, 30, 0x04,  1, "V", "R")},
    {364, T("STDEVPA",          1, 30, 0x04,  1, "V", "R")},
    {365, T("VARPA",            1, 30, 0x04,  1, "V", "R")},
    {366, T("STDEVA",           1, 30, 0x04,  1, "V", "R")},
    {367, T("VARA",             1, 30, 0x04,  1, "V", "R")},
    {368, T("BAHTTEXT",         1,  1, 0x02,  1, "V", "V")},
    {369, T("THAIDAYOFWEEK",    1,  1, 0x02,  1, "V", "V")},
    {370, T("THAIDIGIT",        1,  1, 0x02,  1, "V", "V")},
    {371, T("THAIMONTHOFYEAR",  1,  1, 0x02,  1, "V", "V")},
    {372, T("THAINUMSOUND",     1,  1, 0x02,  1, "V", "V")},
    {373, T("THAINUMSTRING",    1,  1, 0x02,  1, "V", "V")},
    {374, T("THAISTRINGLENGTH", 1,  1, 0x02,  1, "V", "V")},
    {375, T("ISTHAIDIGIT",      1,  1, 0x02,  1, "V", "V")},
    {376, T("ROUNDBAHTDOWN",    1,  1, 0x02,  1, "V", "V")},
    {377, T("ROUNDBAHTUP",      1,  1, 0x02,  1, "V", "V")},
    {378, T("THAIYEAR",         1,  1, 0x02,  1, "V", "V")},
    {379, T("RTD",              2,  5, 0x04,  1, "V", "V")},
};
#undef T

static std::map<int, std::string>
tAttrNames = {
    {0x00, "Skip??"}, // seen in SAMPLES.XLS which shipped with Excel 5.0
    {0x01, "Volatile"},
    {0x02, "If"},
    {0x04, "Choose"},
    {0x08, "Skip"},
    {0x10, "Sum"},
    {0x20, "Assign"},
    {0x40, "Space"},
    {0x41, "SpaceVolatile"},
};

static std::set<int>
error_opcodes = {0x07, 0x08, 0x0A, 0x0B, 0x1C, 0x1D, 0x2F};

// tRangeFuncs = (min, max, min, max, min, max)
// tIsectFuncs = (max, min, max, min, max, min)


/*
def do_box_funcs(box_funcs, boxa, boxb):
    return tuple([
        func(numa, numb)
        for func, numa, numb in zip(box_funcs, boxa.coords, boxb.coords)
        ])
*/

inline
std::tuple<int, int, int, int>
adjust_cell_addr_biff8(int rowval, int colval, int reldelta, int browx, int bcolx) {
    int row_rel = (colval >> 15) & 1;
    int col_rel = (colval >> 14) & 1;
    int rowx = rowval;
    int colx = colval & 0xff;
    if (reldelta) {
        if (row_rel and rowx >= 32768) {
            rowx -= 65536;
        }
        if (col_rel and colx >= 128) {
            colx -= 256;
        }
    }
    else {
        if (row_rel) {
            rowx -= browx;
        }
        if (col_rel) {
            colx -= bcolx;
        }
    }
    return std::make_tuple(rowx, colx, row_rel, col_rel);
}

inline
std::tuple<int, int, int, int>
adjust_cell_addr_biff_le7(int rowval, int colval, int reldelta, int browx, int bcolx) {
    int row_rel = (rowval >> 15) & 1;
    int col_rel = (rowval >> 14) & 1;
    int rowx = rowval & 0x3fff;
    int colx = colval;
    if (reldelta) {
        if (row_rel and rowx >= 8192) {
            rowx -= 16384;
        }
        if (col_rel and colx >= 128) {
            colx -= 256;
        }
    }
    else {
        if (row_rel) {
            rowx -= browx;
        }
        if (col_rel) {
            colx -= bcolx;
        }
    }
    return std::make_tuple(rowx, colx, row_rel, col_rel);
}

inline
std::tuple<int, int, int, int>
get_cell_addr(std::vector<uint8_t> data, int pos, int bv, int reldelta, int browx, int bcolx) {
    if (bv >= 80) {
        int rowval = utils::as_uint16(data, pos);
        int colval = utils::as_uint16(data, pos+2);
        // print "    rv=%04xh cv=%04xh" % (rowval, colval)
        return adjust_cell_addr_biff8(rowval, colval, reldelta, browx, bcolx);
    }
    else {
        int rowval = utils::as_uint16(data, pos);
        int colval = utils::as_uint8(data, pos+2);
        // print "    rv=%04xh cv=%04xh" % (rowval, colval)
        return adjust_cell_addr_biff_le7(
                    rowval, colval, reldelta, browx, bcolx);
    }
}

/*
def get_cell_range_addr(data, pos, bv, reldelta, browx=None, bcolx=None):
    if bv >= 80:
        row1val, row2val, col1val, col2val = unpack("<HHHH", data[pos:pos+8])
        // print "    rv=%04xh cv=%04xh" % (row1val, col1val)
        // print "    rv=%04xh cv=%04xh" % (row2val, col2val)
        res1 = adjust_cell_addr_biff8(row1val, col1val, reldelta, browx, bcolx)
        res2 = adjust_cell_addr_biff8(row2val, col2val, reldelta, browx, bcolx)
        return res1, res2
    else:
        row1val, row2val, col1val, col2val = unpack("<HHBB", data[pos:pos+6])
        // print "    rv=%04xh cv=%04xh" % (row1val, col1val)
        // print "    rv=%04xh cv=%04xh" % (row2val, col2val)
        res1 = adjust_cell_addr_biff_le7(
                    row1val, col1val, reldelta, browx, bcolx)
        res2 = adjust_cell_addr_biff_le7(
                    row2val, col2val, reldelta, browx, bcolx)
        return res1, res2

def get_externsheet_local_range(bk, refx, blah=0):
    try:
        info = bk._externsheet_info[refx]
    except IndexError:
        print("!!! get_externsheet_local_range: refx=%d, not in range(%d)" \
            % (refx, len(bk._externsheet_info)), file=bk.logfile)
        return (-101, -101)
    ref_recordx, ref_first_sheetx, ref_last_sheetx = info
    if ref_recordx == bk._supbook_addins_inx:
        if blah:
            print("/// get_externsheet_local_range(refx=%d) -> addins %r" % (refx, info), file=bk.logfile)
        assert ref_first_sheetx == 0xFFFE == ref_last_sheetx
        return (-5, -5)
    if ref_recordx != bk._supbook_locals_inx:
        if blah:
            print("/// get_externsheet_local_range(refx=%d) -> external %r" % (refx, info), file=bk.logfile)
        return (-4, -4) // external reference
    if ref_first_sheetx == 0xFFFE == ref_last_sheetx:
        if blah:
            print("/// get_externsheet_local_range(refx=%d) -> unspecified sheet %r" % (refx, info), file=bk.logfile)
        return (-1, -1) // internal reference, any sheet
    if ref_first_sheetx == 0xFFFF == ref_last_sheetx:
        if blah:
            print("/// get_externsheet_local_range(refx=%d) -> deleted sheet(s)" % (refx, ), file=bk.logfile)
        return (-2, -2) // internal reference, deleted sheet(s)
    nsheets = len(bk._all_sheets_map)
    if not(0 <= ref_first_sheetx <= ref_last_sheetx < nsheets):
        if blah:
            print("/// get_externsheet_local_range(refx=%d) -> %r" % (refx, info), file=bk.logfile)
            print("--- first/last sheet not in range(%d)" % nsheets, file=bk.logfile)
        return (-102, -102) // stuffed up somewhere :-(
    xlrd_sheetx1 = bk._all_sheets_map[ref_first_sheetx]
    xlrd_sheetx2 = bk._all_sheets_map[ref_last_sheetx]
    if not(0 <= xlrd_sheetx1 <= xlrd_sheetx2):
        return (-3, -3) // internal reference, but to a macro sheet
    return xlrd_sheetx1, xlrd_sheetx2

def get_externsheet_local_range_b57(
        bk, raw_extshtx, ref_first_sheetx, ref_last_sheetx, blah=0):
    if raw_extshtx > 0:
        if blah:
            print("/// get_externsheet_local_range_b57(raw_extshtx=%d) -> external" % raw_extshtx, file=bk.logfile)
        return (-4, -4) // external reference
    if ref_first_sheetx == -1 and ref_last_sheetx == -1:
        return (-2, -2) // internal reference, deleted sheet(s)
    nsheets = len(bk._all_sheets_map)
    if not(0 <= ref_first_sheetx <= ref_last_sheetx < nsheets):
        if blah:
            print("/// get_externsheet_local_range_b57(%d, %d, %d) -> ???" \
                % (raw_extshtx, ref_first_sheetx, ref_last_sheetx), file=bk.logfile)
            print("--- first/last sheet not in range(%d)" % nsheets, file=bk.logfile)
        return (-103, -103) // stuffed up somewhere :-(
    xlrd_sheetx1 = bk._all_sheets_map[ref_first_sheetx]
    xlrd_sheetx2 = bk._all_sheets_map[ref_last_sheetx]
    if not(0 <= xlrd_sheetx1 <= xlrd_sheetx2):
        return (-3, -3) // internal reference, but to a macro sheet
    return xlrd_sheetx1, xlrd_sheetx2

class FormulaError(Exception):
    pass


////
// Used in evaluating formulas.
// The following table describes the kinds and how their values
// are represented.</p>
//
// <table border="1" cellpadding="7">
// <tr>
// <th>Kind symbol</th>
// <th>Kind number</th>
// <th>Value representation</th>
// </tr>
// <tr>
// <td>oBOOL</td>
// <td align="center">3</td>
// <td>integer: 0 => False; 1 => True</td>
// </tr>
// <tr>
// <td>oERR</td>
// <td align="center">4</td>
// <td>None, or an int error code (same as XL_CELL_ERROR in the Cell class).
// </td>
// </tr>
// <tr>
// <td>oMSNG</td>
// <td align="center">5</td>
// <td>Used by Excel as a placeholder for a missing (not supplied) function
// argument. Should *not* appear as a final formula result. Value is None.</td>
// </tr>
// <tr>
// <td>oNUM</td>
// <td align="center">2</td>
// <td>A float. Note that there is no way of distinguishing dates.</td>
// </tr>
// <tr>
// <td>oREF</td>
// <td align="center">-1</td>
// <td>The value is either None or a non-empty list of
// absolute Ref3D instances.<br>
// </td>
// </tr>
// <tr>
// <td>oREL</td>
// <td align="center">-2</td>
// <td>The value is None or a non-empty list of
// fully or partially relative Ref3D instances.
// </td>
// </tr>
// <tr>
// <td>oSTRG</td>
// <td align="center">1</td>
// <td>A Unicode string.</td>
// </tr>
// <tr>
// <td>oUNK</td>
// <td align="center">0</td>
// <td>The kind is unknown or ambiguous. The value is None</td>
// </tr>
// </table>
//<p></p>

*/

class Operand
{
public:
    ////
    // None means that the actual value of the operand is a variable
    // (depends on cell data), not a constant.
    utils::any value;
    ////
    // oUNK means that the kind of operand is not known unambiguously.
    int kind = oUNK;
    ////
    // The reconstituted text of the original formula. Function names will be
    // in English irrespective of the original language, which doesn't seem
    // to be recorded anywhere. The separator is ",", not ";" or whatever else
    // might be more appropriate for the end-user's locale; patches welcome.
    std::string text = "?";
    int rank = 0;

    inline
    Operand(int akind=-1, utils::any avalue=nullptr, int arank=0, std::string atext="?")
    {
        if (akind != -1) {
            this->kind = akind;
        }
        this->value = avalue;
        this->rank = arank;
        // rank is an internal gizmo (operator precedence);
        // it's used in reconstructing formula text.
        this->text = atext;
    };

    inline
    std::string
    repr() {
        auto kind_text = utils::getelse(okind_dict, this->kind, "?Unknown kind?");
        return strutil::format(
            "Operand(kind=%s, value=%s, text=%s)",
            kind_text, this->value, this->text
        );
    };
};

////
// <p>Represents an absolute or relative 3-dimensional reference to a box
// of one or more cells.<br />
// -- New in version 0.6.0
// </p>
//
// <p>The <i>coords</i> attribute is a tuple of the form:<br />
// (shtxlo, shtxhi, rowxlo, rowxhi, colxlo, colxhi)<br />
// where 0 <= thingxlo <= thingx < thingxhi.<br />
// Note that it is quite possible to have thingx > nthings; for example
// Print_Titles could have colxhi == 256 and/or rowxhi == 65536
// irrespective of how many columns/rows are actually used in the worksheet.
// The caller will need to decide how to handle this situation.
// Keyword: IndexError :-)
// </p>
//
// <p>The components of the coords attribute are also available as individual
// attributes: shtxlo, shtxhi, rowxlo, rowxhi, colxlo, and colxhi.</p>
//
// <p>The <i>relflags</i> attribute is a 6-tuple of flags which indicate whether
// the corresponding (sheet|row|col)(lo|hi) is relative (1) or absolute (0).<br>
// Note that there is necessarily no information available as to what cell(s)
// the reference could possibly be relative to. The caller must decide what if
// any use to make of oREL operands. Note also that a partially relative
// reference may well be a typo.
// For example, define name A1Z10 as $a$1:$z10 (missing $ after z)
// while the cursor is on cell Sheet3!A27.<br>
// The resulting Ref3D instance will have coords = (2, 3, 0, -16, 0, 26)
// and relflags = (0, 0, 0, 1, 0, 0).<br>
// So far, only one possibility of a sheet-relative component in
// a reference has been noticed: a 2D reference located in the "current sheet".
// <br /> This will appear as coords = (0, 1, ...) and relflags = (1, 1, ...).
class Ref3D
{
public:
    int shtxlo;
    int shtxhi;
    int rowxlo;
    int rowxhi;
    int colxlo;
    int colxhi;
    std::tuple<int, int, int, int, int, int> coords;
    std::tuple<int, int, int, int, int, int> relflags;

    inline
    Ref3D(std::tuple<int, int, int, int, int, int, int, int, int, int, int, int>atuple)
    {
        this->coords = {std::get<0>(atuple), std::get<1>(atuple), std::get<2>(atuple),
                        std::get<3>(atuple), std::get<4>(atuple), std::get<5>(atuple)};
        this->relflags = {std::get<6>(atuple), std::get<7>(atuple), std::get<8>(atuple),
                          std::get<9>(atuple), std::get<10>(atuple), std::get<11>(atuple)};
        this->shtxlo = std::get<0>(atuple);
        this->shtxhi = std::get<1>(atuple);
        this->rowxlo = std::get<2>(atuple);
        this->rowxhi = std::get<3>(atuple);
        this->colxlo = std::get<4>(atuple);
        this->colxhi = std::get<5>(atuple);
    };

    // def __repr__(self):
    //     if not self.relflags or self.relflags == (0, 0, 0, 0, 0, 0):
    //         return "Ref3D(coords=%r)" % (self.coords, )
    //     else:
    //         return "Ref3D(coords=%r, relflags=%r)" \
    //             % (self.coords, self.relflags)
};

/*
tAdd = 0x03
tSub = 0x04
tMul = 0x05
tDiv = 0x06
tPower = 0x07
tConcat = 0x08
tLT, tLE, tEQ, tGE, tGT, tNE = range(0x09, 0x0F)

import operator as opr

def nop(x):
    return x

def _opr_pow(x, y): return x ** y

def _opr_lt(x, y): return x <  y
def _opr_le(x, y): return x <= y
def _opr_eq(x, y): return x == y
def _opr_ge(x, y): return x >= y
def _opr_gt(x, y): return x >  y
def _opr_ne(x, y): return x != y

def num2strg(num):
    """Attempt to emulate Excel's default conversion
       from number to string.
    """
    s = str(num)
    if s.endswith(".0"):
        s = s[:-2]
    return s

_arith_argdict = {oNUM: nop,     oSTRG: float}
_cmp_argdict =   {oNUM: nop,     oSTRG: nop}
// Seems no conversions done on relops; in Excel, "1" > 9 produces TRUE.
_strg_argdict =  {oNUM:num2strg, oSTRG:nop}
binop_rules = {
    tAdd:   (_arith_argdict, oNUM, opr.add,  30, '+'),
    tSub:   (_arith_argdict, oNUM, opr.sub,  30, '-'),
    tMul:   (_arith_argdict, oNUM, opr.mul,  40, '*'),
    tDiv:   (_arith_argdict, oNUM, opr.truediv,  40, '/'),
    tPower: (_arith_argdict, oNUM, _opr_pow, 50, '^',),
    tConcat:(_strg_argdict, oSTRG, opr.add,  20, '&'),
    tLT:    (_cmp_argdict, oBOOL, _opr_lt,   10, '<'),
    tLE:    (_cmp_argdict, oBOOL, _opr_le,   10, '<='),
    tEQ:    (_cmp_argdict, oBOOL, _opr_eq,   10, '='),
    tGE:    (_cmp_argdict, oBOOL, _opr_ge,   10, '>='),
    tGT:    (_cmp_argdict, oBOOL, _opr_gt,   10, '>'),
    tNE:    (_cmp_argdict, oBOOL, _opr_ne,   10, '<>'),
    }

unop_rules = {
    0x13: (lambda x: -x,        70, '-', ''), // unary minus
    0x12: (lambda x: x,         70, '+', ''), // unary plus
    0x14: (lambda x: x / 100.0, 60, '',  '%'),// percent
    }

LEAF_RANK = 90
FUNC_RANK = 90

STACK_ALARM_LEVEL = 5
STACK_PANIC_LEVEL = 10

def evaluate_name_formula(bk, nobj, namex, blah=0, level=0):
    if level > STACK_ALARM_LEVEL:
        blah = 1
    data = nobj.raw_formula
    fmlalen = nobj.basic_formula_len
    bv = bk.biff_version
    reldelta = 1 // All defined name formulas use "Method B" [OOo docs]
    if blah:
        print("::: evaluate_name_formula %r %r %d %d %r level=%d" \
            % (namex, nobj.name, fmlalen, bv, data, level), file=bk.logfile)
        hex_char_dump(data, 0, fmlalen, fout=bk.logfile)
    if level > STACK_PANIC_LEVEL:
        raise XLRDError("Excessive indirect references in NAME formula")
    sztab = szdict[bv]
    pos = 0
    stack = []
    any_rel = 0
    any_err = 0
    any_external = 0
    unk_opnd = Operand(oUNK, None)
    error_opnd = Operand(oERR, None)
    spush = stack.append

    def do_binop(opcd, stk):
        assert len(stk) >= 2
        bop = stk.pop()
        aop = stk.pop()
        argdict, result_kind, func, rank, sym = binop_rules[opcd]
        otext = ''.join([
            '('[:aop.rank < rank],
            aop.text,
            ')'[:aop.rank < rank],
            sym,
            '('[:bop.rank < rank],
            bop.text,
            ')'[:bop.rank < rank],
            ])
        resop = Operand(result_kind, None, rank, otext)
        try:
            bconv = argdict[bop.kind]
            aconv = argdict[aop.kind]
        except KeyError:
            stk.append(resop)
            return
        if bop.value is None or aop.value is None:
            stk.append(resop)
            return
        bval = bconv(bop.value)
        aval = aconv(aop.value)
        result = func(aval, bval)
        if result_kind == oBOOL:
            result = 1 if result else 0
        resop.value = result
        stk.append(resop)

    def do_unaryop(opcode, result_kind, stk):
        assert len(stk) >= 1
        aop = stk.pop()
        val = aop.value
        func, rank, sym1, sym2 = unop_rules[opcode]
        otext = ''.join([
            sym1,
            '('[:aop.rank < rank],
            aop.text,
            ')'[:aop.rank < rank],
            sym2,
            ])
        if val is not None:
            val = func(val)
        stk.append(Operand(result_kind, val, rank, otext))

    def not_in_name_formula(op_arg, oname_arg):
        msg = "ERROR *** Token 0x%02x (%s) found in NAME formula" \
              % (op_arg, oname_arg)
        raise FormulaError(msg)

    if fmlalen == 0:
        stack = [unk_opnd]

    while 0 <= pos < fmlalen:
        op = BYTES_ORD(data[pos])
        opcode = op & 0x1f
        optype = (op & 0x60) >> 5
        if optype:
            opx = opcode + 32
        else:
            opx = opcode
        oname = onames[opx] // + [" RVA"][optype]
        sz = sztab[opx]
        if blah:
            print("Pos:%d Op:0x%02x Name:t%s Sz:%d opcode:%02xh optype:%02xh" \
                % (pos, op, oname, sz, opcode, optype), file=bk.logfile)
            print("Stack =", stack, file=bk.logfile)
        if sz == -2:
            msg = 'ERROR *** Unexpected token 0x%02x ("%s"); biff_version=%d' \
                % (op, oname, bv)
            raise FormulaError(msg)
        if not optype:
            if 0x00 <= opcode <= 0x02: // unk_opnd, tExp, tTbl
                not_in_name_formula(op, oname)
            elif 0x03 <= opcode <= 0x0E:
                // Add, Sub, Mul, Div, Power
                // tConcat
                // tLT, ..., tNE
                do_binop(opcode, stack)
            elif opcode == 0x0F: // tIsect
                if blah: print("tIsect pre", stack, file=bk.logfile)
                assert len(stack) >= 2
                bop = stack.pop()
                aop = stack.pop()
                sym = ' '
                rank = 80 //////////////////// check //////////////
                otext = ''.join([
                    '('[:aop.rank < rank],
                    aop.text,
                    ')'[:aop.rank < rank],
                    sym,
                    '('[:bop.rank < rank],
                    bop.text,
                    ')'[:bop.rank < rank],
                    ])
                res = Operand(oREF)
                res.text = otext
                if bop.kind == oERR or aop.kind == oERR:
                    res.kind = oERR
                elif bop.kind == oUNK or aop.kind == oUNK:
                    // This can happen with undefined
                    // (go search in the current sheet) labels.
                    // For example =Bob Sales
                    // Each label gets a NAME record with an empty formula (!)
                    // Evaluation of the tName token classifies it as oUNK
                    // res.kind = oREF
                    pass
                elif bop.kind == oREF == aop.kind:
                    if aop.value is not None and bop.value is not None:
                        assert len(aop.value) == 1
                        assert len(bop.value) == 1
                        coords = do_box_funcs(
                            tIsectFuncs, aop.value[0], bop.value[0])
                        res.value = [Ref3D(coords)]
                elif bop.kind == oREL == aop.kind:
                    res.kind = oREL
                    if aop.value is not None and bop.value is not None:
                        assert len(aop.value) == 1
                        assert len(bop.value) == 1
                        coords = do_box_funcs(
                            tIsectFuncs, aop.value[0], bop.value[0])
                        relfa = aop.value[0].relflags
                        relfb = bop.value[0].relflags
                        if relfa == relfb:
                            res.value = [Ref3D(coords + relfa)]
                else:
                    pass
                spush(res)
                if blah: print("tIsect post", stack, file=bk.logfile)
            elif opcode == 0x10: // tList
                if blah: print("tList pre", stack, file=bk.logfile)
                assert len(stack) >= 2
                bop = stack.pop()
                aop = stack.pop()
                sym = ','
                rank = 80 //////////////////// check //////////////
                otext = ''.join([
                    '('[:aop.rank < rank],
                    aop.text,
                    ')'[:aop.rank < rank],
                    sym,
                    '('[:bop.rank < rank],
                    bop.text,
                    ')'[:bop.rank < rank],
                    ])
                res = Operand(oREF, None, rank, otext)
                if bop.kind == oERR or aop.kind == oERR:
                    res.kind = oERR
                elif bop.kind in (oREF, oREL) and aop.kind in (oREF, oREL):
                    res.kind = oREF
                    if aop.kind == oREL or bop.kind == oREL:
                        res.kind = oREL
                    if aop.value is not None and bop.value is not None:
                        assert len(aop.value) >= 1
                        assert len(bop.value) == 1
                        res.value = aop.value + bop.value
                else:
                    pass
                spush(res)
                if blah: print("tList post", stack, file=bk.logfile)
            elif opcode == 0x11: // tRange
                if blah: print("tRange pre", stack, file=bk.logfile)
                assert len(stack) >= 2
                bop = stack.pop()
                aop = stack.pop()
                sym = ':'
                rank = 80 //////////////////// check //////////////
                otext = ''.join([
                    '('[:aop.rank < rank],
                    aop.text,
                    ')'[:aop.rank < rank],
                    sym,
                    '('[:bop.rank < rank],
                    bop.text,
                    ')'[:bop.rank < rank],
                    ])
                res = Operand(oREF, None, rank, otext)
                if bop.kind == oERR or aop.kind == oERR:
                    res = oERR
                elif bop.kind == oREF == aop.kind:
                    if aop.value is not None and bop.value is not None:
                        assert len(aop.value) == 1
                        assert len(bop.value) == 1
                        coords = do_box_funcs(
                            tRangeFuncs, aop.value[0], bop.value[0])
                        res.value = [Ref3D(coords)]
                elif bop.kind == oREL == aop.kind:
                    res.kind = oREL
                    if aop.value is not None and bop.value is not None:
                        assert len(aop.value) == 1
                        assert len(bop.value) == 1
                        coords = do_box_funcs(
                            tRangeFuncs, aop.value[0], bop.value[0])
                        relfa = aop.value[0].relflags
                        relfb = bop.value[0].relflags
                        if relfa == relfb:
                            res.value = [Ref3D(coords + relfa)]
                else:
                    pass
                spush(res)
                if blah: print("tRange post", stack, file=bk.logfile)
            elif 0x12 <= opcode <= 0x14: // tUplus, tUminus, tPercent
                do_unaryop(opcode, oNUM, stack)
            elif opcode == 0x15: // tParen
                // source cosmetics
                pass
            elif opcode == 0x16: // tMissArg
                spush(Operand(oMSNG, None, LEAF_RANK, ''))
            elif opcode == 0x17: // tStr
                if bv <= 70:
                    strg, newpos = unpack_string_update_pos(
                                        data, pos+1, bk.encoding, lenlen=1)
                else:
                    strg, newpos = unpack_unicode_update_pos(
                                        data, pos+1, lenlen=1)
                sz = newpos - pos
                if blah: print("   sz=%d strg=%r" % (sz, strg), file=bk.logfile)
                text = '"' + strg.replace('"', '""') + '"'
                spush(Operand(oSTRG, strg, LEAF_RANK, text))
            elif opcode == 0x18: // tExtended
                // new with BIFF 8
                assert bv >= 80
                // not in OOo docs
                raise FormulaError("tExtended token not implemented")
            elif opcode == 0x19: // tAttr
                subop, nc = unpack("<BH", data[pos+1:pos+4])
                subname = tAttrNames.get(subop, "??Unknown??")
                if subop == 0x04: // Choose
                    sz = nc * 2 + 6
                elif subop == 0x10: // Sum (single arg)
                    sz = 4
                    if blah: print("tAttrSum", stack, file=bk.logfile)
                    assert len(stack) >= 1
                    aop = stack[-1]
                    otext = 'SUM(%s)' % aop.text
                    stack[-1] = Operand(oNUM, None, FUNC_RANK, otext)
                else:
                    sz = 4
                if blah:
                    print("   subop=%02xh subname=t%s sz=%d nc=%02xh" \
                        % (subop, subname, sz, nc), file=bk.logfile)
            elif 0x1A <= opcode <= 0x1B: // tSheet, tEndSheet
                assert bv < 50
                raise FormulaError("tSheet & tEndsheet tokens not implemented")
            elif 0x1C <= opcode <= 0x1F: // tErr, tBool, tInt, tNum
                inx = opcode - 0x1C
                nb = [1, 1, 2, 8][inx]
                kind = [oERR, oBOOL, oNUM, oNUM][inx]
                value, = unpack("<" + "BBHd"[inx], data[pos+1:pos+1+nb])
                if inx == 2: // tInt
                    value = float(value)
                    text = str(value)
                elif inx == 3: // tNum
                    text = str(value)
                elif inx == 1: // tBool
                    text = ('FALSE', 'TRUE')[value]
                else:
                    text = '"' +error_text_from_code[value] + '"'
                spush(Operand(kind, value, LEAF_RANK, text))
            else:
                raise FormulaError("Unhandled opcode: 0x%02x" % opcode)
            if sz <= 0:
                raise FormulaError("Size not set for opcode 0x%02x" % opcode)
            pos += sz
            continue
        if opcode == 0x00: // tArray
            spush(unk_opnd)
        elif opcode == 0x01: // tFunc
            nb = 1 + int(bv >= 40)
            funcx = unpack("<" + " BH"[nb], data[pos+1:pos+1+nb])[0]
            func_attrs = func_defs.get(funcx, None)
            if not func_attrs:
                print("*** formula/tFunc unknown FuncID:%d" \
                      % funcx, file=bk.logfile)
                spush(unk_opnd)
            else:
                func_name, nargs = func_attrs[:2]
                if blah:
                    print("    FuncID=%d name=%s nargs=%d" \
                          % (funcx, func_name, nargs), file=bk.logfile)
                assert len(stack) >= nargs
                if nargs:
                    argtext = listsep.join([arg.text for arg in stack[-nargs:]])
                    otext = "%s(%s)" % (func_name, argtext)
                    del stack[-nargs:]
                else:
                    otext = func_name + "()"
                res = Operand(oUNK, None, FUNC_RANK, otext)
                spush(res)
        elif opcode == 0x02: //tFuncVar
            nb = 1 + int(bv >= 40)
            nargs, funcx = unpack("<B" + " BH"[nb], data[pos+1:pos+2+nb])
            prompt, nargs = divmod(nargs, 128)
            macro, funcx = divmod(funcx, 32768)
            if blah:
                print("   FuncID=%d nargs=%d macro=%d prompt=%d" \
                      % (funcx, nargs, macro, prompt), file=bk.logfile)
            func_attrs = func_defs.get(funcx, None)
            if not func_attrs:
                print("*** formula/tFuncVar unknown FuncID:%d" \
                      % funcx, file=bk.logfile)
                spush(unk_opnd)
            else:
                func_name, minargs, maxargs = func_attrs[:3]
                if blah:
                    print("    name: %r, min~max args: %d~%d" \
                        % (func_name, minargs, maxargs), file=bk.logfile)
                assert minargs <= nargs <= maxargs
                assert len(stack) >= nargs
                assert len(stack) >= nargs
                argtext = listsep.join([arg.text for arg in stack[-nargs:]])
                otext = "%s(%s)" % (func_name, argtext)
                res = Operand(oUNK, None, FUNC_RANK, otext)
                if funcx == 1: // IF
                    testarg = stack[-nargs]
                    if testarg.kind not in (oNUM, oBOOL):
                        if blah and testarg.kind != oUNK:
                            print("IF testarg kind?", file=bk.logfile)
                    elif testarg.value not in (0, 1):
                        if blah and testarg.value is not None:
                            print("IF testarg value?", file=bk.logfile)
                    else:
                        if nargs == 2 and not testarg.value:
                            // IF(FALSE, tv) => FALSE
                            res.kind, res.value = oBOOL, 0
                        else:
                            respos = -nargs + 2 - int(testarg.value)
                            chosen = stack[respos]
                            if chosen.kind == oMSNG:
                                res.kind, res.value = oNUM, 0
                            else:
                                res.kind, res.value = chosen.kind, chosen.value
                        if blah:
                            print("$$$$$$ IF => constant", file=bk.logfile)
                elif funcx == 100: // CHOOSE
                    testarg = stack[-nargs]
                    if testarg.kind == oNUM:
                        if 1 <= testarg.value < nargs:
                            chosen = stack[-nargs + int(testarg.value)]
                            if chosen.kind == oMSNG:
                                res.kind, res.value = oNUM, 0
                            else:
                                res.kind, res.value = chosen.kind, chosen.value
                del stack[-nargs:]
                spush(res)
        elif opcode == 0x03: //tName
            tgtnamex = unpack("<H", data[pos+1:pos+3])[0] - 1
            // Only change with BIFF version is number of trailing UNUSED bytes!
            if blah: print("   tgtnamex=%d" % tgtnamex, file=bk.logfile)
            tgtobj = bk.name_obj_list[tgtnamex]
            if not tgtobj.evaluated:
                ////// recursive //////
                evaluate_name_formula(bk, tgtobj, tgtnamex, blah, level+1)
            if tgtobj.macro or tgtobj.binary \
            or tgtobj.any_err:
                if blah:
                    tgtobj.dump(
                        bk.logfile,
                        header="!!! tgtobj has problems!!!",
                        footer="-----------       --------",
                        )
                res = Operand(oUNK, None)
                any_err = any_err or tgtobj.macro or tgtobj.binary or tgtobj.any_err
                any_rel = any_rel or tgtobj.any_rel
            else:
                assert len(tgtobj.stack) == 1
                res = copy.deepcopy(tgtobj.stack[0])
            res.rank = LEAF_RANK
            if tgtobj.scope == -1:
                res.text = tgtobj.name
            else:
                res.text = "%s!%s" \
                           % (bk._sheet_names[tgtobj.scope], tgtobj.name)
            if blah:
                print("    tName: setting text to", repr(res.text), file=bk.logfile)
            spush(res)
        elif opcode == 0x04: // tRef
            // not_in_name_formula(op, oname)
            res = get_cell_addr(data, pos+1, bv, reldelta)
            if blah: print("  ", res, file=bk.logfile)
            rowx, colx, row_rel, col_rel = res
            shx1 = shx2 = 0 ////////////// N.B. relative to the CURRENT SHEET
            any_rel = 1
            coords = (shx1, shx2+1, rowx, rowx+1, colx, colx+1)
            if blah: print("   ", coords, file=bk.logfile)
            res = Operand(oUNK, None)
            if optype == 1:
                relflags = (1, 1, row_rel, row_rel, col_rel, col_rel)
                res = Operand(oREL, [Ref3D(coords + relflags)])
            spush(res)
        elif opcode == 0x05: // tArea
            // not_in_name_formula(op, oname)
            res1, res2 = get_cell_range_addr(data, pos+1, bv, reldelta)
            if blah: print("  ", res1, res2, file=bk.logfile)
            rowx1, colx1, row_rel1, col_rel1 = res1
            rowx2, colx2, row_rel2, col_rel2 = res2
            shx1 = shx2 = 0 ////////////// N.B. relative to the CURRENT SHEET
            any_rel = 1
            coords = (shx1, shx2+1, rowx1, rowx2+1, colx1, colx2+1)
            if blah: print("   ", coords, file=bk.logfile)
            res = Operand(oUNK, None)
            if optype == 1:
                relflags = (1, 1, row_rel1, row_rel2, col_rel1, col_rel2)
                res = Operand(oREL, [Ref3D(coords + relflags)])
            spush(res)
        elif opcode == 0x06: // tMemArea
            not_in_name_formula(op, oname)
        elif opcode == 0x09: // tMemFunc
            nb = unpack("<H", data[pos+1:pos+3])[0]
            if blah: print("  %d bytes of cell ref formula" % nb, file=bk.logfile)
            // no effect on stack
        elif opcode == 0x0C: //tRefN
            not_in_name_formula(op, oname)
            // res = get_cell_addr(data, pos+1, bv, reldelta=1)
            // // note *ALL* tRefN usage has signed offset for relative addresses
            // any_rel = 1
            // if blah: print >> bk.logfile, "   ", res
            // spush(res)
        elif opcode == 0x0D: //tAreaN
            not_in_name_formula(op, oname)
            // res = get_cell_range_addr(data, pos+1, bv, reldelta=1)
            // // note *ALL* tAreaN usage has signed offset for relative addresses
            // any_rel = 1
            // if blah: print >> bk.logfile, "   ", res
        elif opcode == 0x1A: // tRef3d
            if bv >= 80:
                res = get_cell_addr(data, pos+3, bv, reldelta)
                refx = unpack("<H", data[pos+1:pos+3])[0]
                shx1, shx2 = get_externsheet_local_range(bk, refx, blah)
            else:
                res = get_cell_addr(data, pos+15, bv, reldelta)
                raw_extshtx, raw_shx1, raw_shx2 = \
                             unpack("<hxxxxxxxxhh", data[pos+1:pos+15])
                if blah:
                    print("tRef3d", raw_extshtx, raw_shx1, raw_shx2, file=bk.logfile)
                shx1, shx2 = get_externsheet_local_range_b57(
                                bk, raw_extshtx, raw_shx1, raw_shx2, blah)
            rowx, colx, row_rel, col_rel = res
            is_rel = row_rel or col_rel
            any_rel = any_rel or is_rel
            coords = (shx1, shx2+1, rowx, rowx+1, colx, colx+1)
            any_err |= shx1 < -1
            if blah: print("   ", coords, file=bk.logfile)
            res = Operand(oUNK, None)
            if is_rel:
                relflags = (0, 0, row_rel, row_rel, col_rel, col_rel)
                ref3d = Ref3D(coords + relflags)
                res.kind = oREL
                res.text = rangename3drel(bk, ref3d, r1c1=1)
            else:
                ref3d = Ref3D(coords)
                res.kind = oREF
                res.text = rangename3d(bk, ref3d)
            res.rank = LEAF_RANK
            if optype == 1:
                res.value = [ref3d]
            spush(res)
        elif opcode == 0x1B: // tArea3d
            if bv >= 80:
                res1, res2 = get_cell_range_addr(data, pos+3, bv, reldelta)
                refx = unpack("<H", data[pos+1:pos+3])[0]
                shx1, shx2 = get_externsheet_local_range(bk, refx, blah)
            else:
                res1, res2 = get_cell_range_addr(data, pos+15, bv, reldelta)
                raw_extshtx, raw_shx1, raw_shx2 = \
                             unpack("<hxxxxxxxxhh", data[pos+1:pos+15])
                if blah:
                    print("tArea3d", raw_extshtx, raw_shx1, raw_shx2, file=bk.logfile)
                shx1, shx2 = get_externsheet_local_range_b57(
                                bk, raw_extshtx, raw_shx1, raw_shx2, blah)
            any_err |= shx1 < -1
            rowx1, colx1, row_rel1, col_rel1 = res1
            rowx2, colx2, row_rel2, col_rel2 = res2
            is_rel = row_rel1 or col_rel1 or row_rel2 or col_rel2
            any_rel = any_rel or is_rel
            coords = (shx1, shx2+1, rowx1, rowx2+1, colx1, colx2+1)
            if blah: print("   ", coords, file=bk.logfile)
            res = Operand(oUNK, None)
            if is_rel:
                relflags = (0, 0, row_rel1, row_rel2, col_rel1, col_rel2)
                ref3d = Ref3D(coords + relflags)
                res.kind = oREL
                res.text = rangename3drel(bk, ref3d, r1c1=1)
            else:
                ref3d = Ref3D(coords)
                res.kind = oREF
                res.text = rangename3d(bk, ref3d)
            res.rank = LEAF_RANK
            if optype == 1:
                res.value = [ref3d]

            spush(res)
        elif opcode == 0x19: // tNameX
            dodgy = 0
            res = Operand(oUNK, None)
            if bv >= 80:
                refx, tgtnamex = unpack("<HH", data[pos+1:pos+5])
                tgtnamex -= 1
                origrefx = refx
            else:
                refx, tgtnamex = unpack("<hxxxxxxxxH", data[pos+1:pos+13])
                tgtnamex -= 1
                origrefx = refx
                if refx > 0:
                    refx -= 1
                elif refx < 0:
                    refx = -refx - 1
                else:
                    dodgy = 1
            if blah:
                print("   origrefx=%d refx=%d tgtnamex=%d dodgy=%d" \
                    % (origrefx, refx, tgtnamex, dodgy), file=bk.logfile)
            if tgtnamex == namex:
                if blah: print("!!!! Self-referential !!!!", file=bk.logfile)
                dodgy = any_err = 1
            if not dodgy:
                if bv >= 80:
                    shx1, shx2 = get_externsheet_local_range(bk, refx, blah)
                elif origrefx > 0:
                    shx1, shx2 = (-4, -4) // external ref
                else:
                    exty = bk._externsheet_type_b57[refx]
                    if exty == 4: // non-specific sheet in own doc't
                        shx1, shx2 = (-1, -1) // internal, any sheet
                    else:
                        shx1, shx2 = (-666, -666)
            if dodgy or shx1 < -1:
                otext = "<<Name //%d in external(?) file //%d>>" \
                        % (tgtnamex, origrefx)
                res = Operand(oUNK, None, LEAF_RANK, otext)
            else:
                tgtobj = bk.name_obj_list[tgtnamex]
                if not tgtobj.evaluated:
                    ////// recursive //////
                    evaluate_name_formula(bk, tgtobj, tgtnamex, blah, level+1)
                if tgtobj.macro or tgtobj.binary \
                or tgtobj.any_err:
                    if blah:
                        tgtobj.dump(
                            bk.logfile,
                            header="!!! bad tgtobj !!!",
                            footer="------------------",
                            )
                    res = Operand(oUNK, None)
                    any_err = any_err or tgtobj.macro or tgtobj.binary or tgtobj.any_err
                    any_rel = any_rel or tgtobj.any_rel
                else:
                    assert len(tgtobj.stack) == 1
                    res = copy.deepcopy(tgtobj.stack[0])
                res.rank = LEAF_RANK
                if tgtobj.scope == -1:
                    res.text = tgtobj.name
                else:
                    res.text = "%s!%s" \
                               % (bk._sheet_names[tgtobj.scope], tgtobj.name)
                if blah:
                    print("    tNameX: setting text to", repr(res.text), file=bk.logfile)
            spush(res)
        elif opcode in error_opcodes:
            any_err = 1
            spush(error_opnd)
        else:
            if blah:
                print("FORMULA: /// Not handled yet: t" + oname, file=bk.logfile)
            any_err = 1
        if sz <= 0:
            raise FormulaError("Fatal: token size is not positive")
        pos += sz
    any_rel = not not any_rel
    if blah:
        fprintf(bk.logfile, "End of formula. level=%d any_rel=%d any_err=%d stack=%r\n",
            level, not not any_rel, any_err, stack)
        if len(stack) >= 2:
            print("*** Stack has unprocessed args", file=bk.logfile)
        print(file=bk.logfile)
    nobj.stack = stack
    if len(stack) != 1:
        nobj.result = None
    else:
        nobj.result = stack[0]
    nobj.any_rel = any_rel
    nobj.any_err = any_err
    nobj.any_external = any_external
    nobj.evaluated = 1

//////// under construction //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
def decompile_formula(bk, fmla, fmlalen,
    fmlatype=None, browx=None, bcolx=None,
    blah=0, level=0, r1c1=0):
    if level > STACK_ALARM_LEVEL:
        blah = 1
    reldelta = fmlatype in (FMLA_TYPE_SHARED, FMLA_TYPE_NAME, FMLA_TYPE_COND_FMT, FMLA_TYPE_DATA_VAL)
    data = fmla
    bv = bk.biff_version
    if blah:
        print("::: decompile_formula len=%d fmlatype=%r browx=%r bcolx=%r reldelta=%d %r level=%d" \
            % (fmlalen, fmlatype, browx, bcolx, reldelta, data, level), file=bk.logfile)
        hex_char_dump(data, 0, fmlalen, fout=bk.logfile)
    if level > STACK_PANIC_LEVEL:
        raise XLRDError("Excessive indirect references in formula")
    sztab = szdict[bv]
    pos = 0
    stack = []
    any_rel = 0
    any_err = 0
    any_external = 0
    unk_opnd = Operand(oUNK, None)
    error_opnd = Operand(oERR, None)
    spush = stack.append

    def do_binop(opcd, stk):
        assert len(stk) >= 2
        bop = stk.pop()
        aop = stk.pop()
        argdict, result_kind, func, rank, sym = binop_rules[opcd]
        otext = ''.join([
            '('[:aop.rank < rank],
            aop.text,
            ')'[:aop.rank < rank],
            sym,
            '('[:bop.rank < rank],
            bop.text,
            ')'[:bop.rank < rank],
            ])
        resop = Operand(result_kind, None, rank, otext)
        stk.append(resop)

    def do_unaryop(opcode, result_kind, stk):
        assert len(stk) >= 1
        aop = stk.pop()
        func, rank, sym1, sym2 = unop_rules[opcode]
        otext = ''.join([
            sym1,
            '('[:aop.rank < rank],
            aop.text,
            ')'[:aop.rank < rank],
            sym2,
            ])
        stk.append(Operand(result_kind, None, rank, otext))

    def unexpected_opcode(op_arg, oname_arg):
        msg = "ERROR *** Unexpected token 0x%02x (%s) found in formula type %s" \
              % (op_arg, oname_arg, FMLA_TYPEDESCR_MAP[fmlatype])
        print(msg, file=bk.logfile)
        // raise FormulaError(msg)

    if fmlalen == 0:
        stack = [unk_opnd]

    while 0 <= pos < fmlalen:
        op = BYTES_ORD(data[pos])
        opcode = op & 0x1f
        optype = (op & 0x60) >> 5
        if optype:
            opx = opcode + 32
        else:
            opx = opcode
        oname = onames[opx] // + [" RVA"][optype]
        sz = sztab[opx]
        if blah:
            print("Pos:%d Op:0x%02x opname:t%s Sz:%d opcode:%02xh optype:%02xh" \
                % (pos, op, oname, sz, opcode, optype), file=bk.logfile)
            print("Stack =", stack, file=bk.logfile)
        if sz == -2:
            msg = 'ERROR *** Unexpected token 0x%02x ("%s"); biff_version=%d' \
                % (op, oname, bv)
            raise FormulaError(msg)
        if _TOKEN_NOT_ALLOWED(opx, 0) & fmlatype:
            unexpected_opcode(op, oname)
        if not optype:
            if opcode <= 0x01: // tExp
                if bv >= 30:
                    fmt = '<x2H'
                else:
                    fmt = '<xHB'
                assert pos == 0 and fmlalen == sz and not stack
                rowx, colx = unpack(fmt, data)
                text = "SHARED FMLA at rowx=%d colx=%d" % (rowx, colx)
                spush(Operand(oUNK, None, LEAF_RANK, text))
                if not fmlatype & (FMLA_TYPE_CELL | FMLA_TYPE_ARRAY):
                    unexpected_opcode(op, oname)
            elif 0x03 <= opcode <= 0x0E:
                // Add, Sub, Mul, Div, Power
                // tConcat
                // tLT, ..., tNE
                do_binop(opcode, stack)
            elif opcode == 0x0F: // tIsect
                if blah: print("tIsect pre", stack, file=bk.logfile)
                assert len(stack) >= 2
                bop = stack.pop()
                aop = stack.pop()
                sym = ' '
                rank = 80 //////////////////// check //////////////
                otext = ''.join([
                    '('[:aop.rank < rank],
                    aop.text,
                    ')'[:aop.rank < rank],
                    sym,
                    '('[:bop.rank < rank],
                    bop.text,
                    ')'[:bop.rank < rank],
                    ])
                res = Operand(oREF)
                res.text = otext
                if bop.kind == oERR or aop.kind == oERR:
                    res.kind = oERR
                elif bop.kind == oUNK or aop.kind == oUNK:
                    // This can happen with undefined
                    // (go search in the current sheet) labels.
                    // For example =Bob Sales
                    // Each label gets a NAME record with an empty formula (!)
                    // Evaluation of the tName token classifies it as oUNK
                    // res.kind = oREF
                    pass
                elif bop.kind == oREF == aop.kind:
                    pass
                elif bop.kind == oREL == aop.kind:
                    res.kind = oREL
                else:
                    pass
                spush(res)
                if blah: print("tIsect post", stack, file=bk.logfile)
            elif opcode == 0x10: // tList
                if blah: print("tList pre", stack, file=bk.logfile)
                assert len(stack) >= 2
                bop = stack.pop()
                aop = stack.pop()
                sym = ','
                rank = 80 //////////////////// check //////////////
                otext = ''.join([
                    '('[:aop.rank < rank],
                    aop.text,
                    ')'[:aop.rank < rank],
                    sym,
                    '('[:bop.rank < rank],
                    bop.text,
                    ')'[:bop.rank < rank],
                    ])
                res = Operand(oREF, None, rank, otext)
                if bop.kind == oERR or aop.kind == oERR:
                    res.kind = oERR
                elif bop.kind in (oREF, oREL) and aop.kind in (oREF, oREL):
                    res.kind = oREF
                    if aop.kind == oREL or bop.kind == oREL:
                        res.kind = oREL
                else:
                    pass
                spush(res)
                if blah: print("tList post", stack, file=bk.logfile)
            elif opcode == 0x11: // tRange
                if blah: print("tRange pre", stack, file=bk.logfile)
                assert len(stack) >= 2
                bop = stack.pop()
                aop = stack.pop()
                sym = ':'
                rank = 80 //////////////////// check //////////////
                otext = ''.join([
                    '('[:aop.rank < rank],
                    aop.text,
                    ')'[:aop.rank < rank],
                    sym,
                    '('[:bop.rank < rank],
                    bop.text,
                    ')'[:bop.rank < rank],
                    ])
                res = Operand(oREF, None, rank, otext)
                if bop.kind == oERR or aop.kind == oERR:
                    res = oERR
                elif bop.kind == oREF == aop.kind:
                    pass
                else:
                    pass
                spush(res)
                if blah: print("tRange post", stack, file=bk.logfile)
            elif 0x12 <= opcode <= 0x14: // tUplus, tUminus, tPercent
                do_unaryop(opcode, oNUM, stack)
            elif opcode == 0x15: // tParen
                // source cosmetics
                pass
            elif opcode == 0x16: // tMissArg
                spush(Operand(oMSNG, None, LEAF_RANK, ''))
            elif opcode == 0x17: // tStr
                if bv <= 70:
                    strg, newpos = unpack_string_update_pos(
                                        data, pos+1, bk.encoding, lenlen=1)
                else:
                    strg, newpos = unpack_unicode_update_pos(
                                        data, pos+1, lenlen=1)
                sz = newpos - pos
                if blah: print("   sz=%d strg=%r" % (sz, strg), file=bk.logfile)
                text = '"' + strg.replace('"', '""') + '"'
                spush(Operand(oSTRG, None, LEAF_RANK, text))
            elif opcode == 0x18: // tExtended
                // new with BIFF 8
                assert bv >= 80
                // not in OOo docs, don't even know how to determine its length
                raise FormulaError("tExtended token not implemented")
            elif opcode == 0x19: // tAttr
                subop, nc = unpack("<BH", data[pos+1:pos+4])
                subname = tAttrNames.get(subop, "??Unknown??")
                if subop == 0x04: // Choose
                    sz = nc * 2 + 6
                elif subop == 0x10: // Sum (single arg)
                    sz = 4
                    if blah: print("tAttrSum", stack, file=bk.logfile)
                    assert len(stack) >= 1
                    aop = stack[-1]
                    otext = 'SUM(%s)' % aop.text
                    stack[-1] = Operand(oNUM, None, FUNC_RANK, otext)
                else:
                    sz = 4
                if blah:
                    print("   subop=%02xh subname=t%s sz=%d nc=%02xh" \
                        % (subop, subname, sz, nc), file=bk.logfile)
            elif 0x1A <= opcode <= 0x1B: // tSheet, tEndSheet
                assert bv < 50
                raise FormulaError("tSheet & tEndsheet tokens not implemented")
            elif 0x1C <= opcode <= 0x1F: // tErr, tBool, tInt, tNum
                inx = opcode - 0x1C
                nb = [1, 1, 2, 8][inx]
                kind = [oERR, oBOOL, oNUM, oNUM][inx]
                value, = unpack("<" + "BBHd"[inx], data[pos+1:pos+1+nb])
                if inx == 2: // tInt
                    value = float(value)
                    text = str(value)
                elif inx == 3: // tNum
                    text = str(value)
                elif inx == 1: // tBool
                    text = ('FALSE', 'TRUE')[value]
                else:
                    text = '"' +error_text_from_code[value] + '"'
                spush(Operand(kind, None, LEAF_RANK, text))
            else:
                raise FormulaError("Unhandled opcode: 0x%02x" % opcode)
            if sz <= 0:
                raise FormulaError("Size not set for opcode 0x%02x" % opcode)
            pos += sz
            continue
        if opcode == 0x00: // tArray
            spush(unk_opnd)
        elif opcode == 0x01: // tFunc
            nb = 1 + int(bv >= 40)
            funcx = unpack("<" + " BH"[nb], data[pos+1:pos+1+nb])[0]
            func_attrs = func_defs.get(funcx, None)
            if not func_attrs:
                print("*** formula/tFunc unknown FuncID:%d" % funcx, file=bk.logfile)
                spush(unk_opnd)
            else:
                func_name, nargs = func_attrs[:2]
                if blah:
                    print("    FuncID=%d name=%s nargs=%d" \
                          % (funcx, func_name, nargs), file=bk.logfile)
                assert len(stack) >= nargs
                if nargs:
                    argtext = listsep.join([arg.text for arg in stack[-nargs:]])
                    otext = "%s(%s)" % (func_name, argtext)
                    del stack[-nargs:]
                else:
                    otext = func_name + "()"
                res = Operand(oUNK, None, FUNC_RANK, otext)
                spush(res)
        elif opcode == 0x02: //tFuncVar
            nb = 1 + int(bv >= 40)
            nargs, funcx = unpack("<B" + " BH"[nb], data[pos+1:pos+2+nb])
            prompt, nargs = divmod(nargs, 128)
            macro, funcx = divmod(funcx, 32768)
            if blah:
                print("   FuncID=%d nargs=%d macro=%d prompt=%d" \
                      % (funcx, nargs, macro, prompt), file=bk.logfile)
            //////// TODO //////// if funcx == 255: // call add-in function
            if funcx == 255:
                func_attrs = ("CALL_ADDIN", 1, 30)
            else:
                func_attrs = func_defs.get(funcx, None)
            if not func_attrs:
                print("*** formula/tFuncVar unknown FuncID:%d" \
                      % funcx, file=bk.logfile)
                spush(unk_opnd)
            else:
                func_name, minargs, maxargs = func_attrs[:3]
                if blah:
                    print("    name: %r, min~max args: %d~%d" \
                        % (func_name, minargs, maxargs), file=bk.logfile)
                assert minargs <= nargs <= maxargs
                assert len(stack) >= nargs
                assert len(stack) >= nargs
                argtext = listsep.join([arg.text for arg in stack[-nargs:]])
                otext = "%s(%s)" % (func_name, argtext)
                res = Operand(oUNK, None, FUNC_RANK, otext)
                del stack[-nargs:]
                spush(res)
        elif opcode == 0x03: //tName
            tgtnamex = unpack("<H", data[pos+1:pos+3])[0] - 1
            // Only change with BIFF version is number of trailing UNUSED bytes!
            if blah: print("   tgtnamex=%d" % tgtnamex, file=bk.logfile)
            tgtobj = bk.name_obj_list[tgtnamex]
            if tgtobj.scope == -1:
                otext = tgtobj.name
            else:
                otext = "%s!%s" % (bk._sheet_names[tgtobj.scope], tgtobj.name)
            if blah:
                print("    tName: setting text to", repr(otext), file=bk.logfile)
            res = Operand(oUNK, None, LEAF_RANK, otext)
            spush(res)
        elif opcode == 0x04: // tRef
            res = get_cell_addr(data, pos+1, bv, reldelta, browx, bcolx)
            if blah: print("  ", res, file=bk.logfile)
            rowx, colx, row_rel, col_rel = res
            is_rel = row_rel or col_rel
            if is_rel:
                okind = oREL
            else:
                okind = oREF
            otext = cellnamerel(rowx, colx, row_rel, col_rel, browx, bcolx, r1c1)
            res = Operand(okind, None, LEAF_RANK, otext)
            spush(res)
        elif opcode == 0x05: // tArea
            res1, res2 = get_cell_range_addr(
                            data, pos+1, bv, reldelta, browx, bcolx)
            if blah: print("  ", res1, res2, file=bk.logfile)
            rowx1, colx1, row_rel1, col_rel1 = res1
            rowx2, colx2, row_rel2, col_rel2 = res2
            coords = (rowx1, rowx2+1, colx1, colx2+1)
            relflags = (row_rel1, row_rel2, col_rel1, col_rel2)
            if sum(relflags):  // relative
                okind = oREL
            else:
                okind = oREF
            if blah: print("   ", coords, relflags, file=bk.logfile)
            otext = rangename2drel(coords, relflags, browx, bcolx, r1c1)
            res = Operand(okind, None, LEAF_RANK, otext)
            spush(res)
        elif opcode == 0x06: // tMemArea
            not_in_name_formula(op, oname)
        elif opcode == 0x09: // tMemFunc
            nb = unpack("<H", data[pos+1:pos+3])[0]
            if blah: print("  %d bytes of cell ref formula" % nb, file=bk.logfile)
            // no effect on stack
        elif opcode == 0x0C: //tRefN
            res = get_cell_addr(data, pos+1, bv, reldelta, browx, bcolx)
            // note *ALL* tRefN usage has signed offset for relative addresses
            any_rel = 1
            if blah: print("   ", res, file=bk.logfile)
            rowx, colx, row_rel, col_rel = res
            is_rel = row_rel or col_rel
            if is_rel:
                okind = oREL
            else:
                okind = oREF
            otext = cellnamerel(rowx, colx, row_rel, col_rel, browx, bcolx, r1c1)
            res = Operand(okind, None, LEAF_RANK, otext)
            spush(res)
        elif opcode == 0x0D: //tAreaN
            // res = get_cell_range_addr(data, pos+1, bv, reldelta, browx, bcolx)
            // // note *ALL* tAreaN usage has signed offset for relative addresses
            // any_rel = 1
            // if blah: print >> bk.logfile, "   ", res
            res1, res2 = get_cell_range_addr(
                            data, pos+1, bv, reldelta, browx, bcolx)
            if blah: print("  ", res1, res2, file=bk.logfile)
            rowx1, colx1, row_rel1, col_rel1 = res1
            rowx2, colx2, row_rel2, col_rel2 = res2
            coords = (rowx1, rowx2+1, colx1, colx2+1)
            relflags = (row_rel1, row_rel2, col_rel1, col_rel2)
            if sum(relflags):  // relative
                okind = oREL
            else:
                okind = oREF
            if blah: print("   ", coords, relflags, file=bk.logfile)
            otext = rangename2drel(coords, relflags, browx, bcolx, r1c1)
            res = Operand(okind, None, LEAF_RANK, otext)
            spush(res)
        elif opcode == 0x1A: // tRef3d
            if bv >= 80:
                res = get_cell_addr(data, pos+3, bv, reldelta, browx, bcolx)
                refx = unpack("<H", data[pos+1:pos+3])[0]
                shx1, shx2 = get_externsheet_local_range(bk, refx, blah)
            else:
                res = get_cell_addr(data, pos+15, bv, reldelta, browx, bcolx)
                raw_extshtx, raw_shx1, raw_shx2 = \
                             unpack("<hxxxxxxxxhh", data[pos+1:pos+15])
                if blah:
                    print("tRef3d", raw_extshtx, raw_shx1, raw_shx2, file=bk.logfile)
                shx1, shx2 = get_externsheet_local_range_b57(
                                bk, raw_extshtx, raw_shx1, raw_shx2, blah)
            rowx, colx, row_rel, col_rel = res
            is_rel = row_rel or col_rel
            any_rel = any_rel or is_rel
            coords = (shx1, shx2+1, rowx, rowx+1, colx, colx+1)
            any_err |= shx1 < -1
            if blah: print("   ", coords, file=bk.logfile)
            res = Operand(oUNK, None)
            if is_rel:
                relflags = (0, 0, row_rel, row_rel, col_rel, col_rel)
                ref3d = Ref3D(coords + relflags)
                res.kind = oREL
                res.text = rangename3drel(bk, ref3d, browx, bcolx, r1c1)
            else:
                ref3d = Ref3D(coords)
                res.kind = oREF
                res.text = rangename3d(bk, ref3d)
            res.rank = LEAF_RANK
            res.value = None
            spush(res)
        elif opcode == 0x1B: // tArea3d
            if bv >= 80:
                res1, res2 = get_cell_range_addr(data, pos+3, bv, reldelta)
                refx = unpack("<H", data[pos+1:pos+3])[0]
                shx1, shx2 = get_externsheet_local_range(bk, refx, blah)
            else:
                res1, res2 = get_cell_range_addr(data, pos+15, bv, reldelta)
                raw_extshtx, raw_shx1, raw_shx2 = \
                             unpack("<hxxxxxxxxhh", data[pos+1:pos+15])
                if blah:
                    print("tArea3d", raw_extshtx, raw_shx1, raw_shx2, file=bk.logfile)
                shx1, shx2 = get_externsheet_local_range_b57(
                                bk, raw_extshtx, raw_shx1, raw_shx2, blah)
            any_err |= shx1 < -1
            rowx1, colx1, row_rel1, col_rel1 = res1
            rowx2, colx2, row_rel2, col_rel2 = res2
            is_rel = row_rel1 or col_rel1 or row_rel2 or col_rel2
            any_rel = any_rel or is_rel
            coords = (shx1, shx2+1, rowx1, rowx2+1, colx1, colx2+1)
            if blah: print("   ", coords, file=bk.logfile)
            res = Operand(oUNK, None)
            if is_rel:
                relflags = (0, 0, row_rel1, row_rel2, col_rel1, col_rel2)
                ref3d = Ref3D(coords + relflags)
                res.kind = oREL
                res.text = rangename3drel(bk, ref3d, browx, bcolx, r1c1)
            else:
                ref3d = Ref3D(coords)
                res.kind = oREF
                res.text = rangename3d(bk, ref3d)
            res.rank = LEAF_RANK
            spush(res)
        elif opcode == 0x19: // tNameX
            dodgy = 0
            res = Operand(oUNK, None)
            if bv >= 80:
                refx, tgtnamex = unpack("<HH", data[pos+1:pos+5])
                tgtnamex -= 1
                origrefx = refx
            else:
                refx, tgtnamex = unpack("<hxxxxxxxxH", data[pos+1:pos+13])
                tgtnamex -= 1
                origrefx = refx
                if refx > 0:
                    refx -= 1
                elif refx < 0:
                    refx = -refx - 1
                else:
                    dodgy = 1
            if blah:
                print("   origrefx=%d refx=%d tgtnamex=%d dodgy=%d" \
                    % (origrefx, refx, tgtnamex, dodgy), file=bk.logfile)
            // if tgtnamex == namex:
            //     if blah: print >> bk.logfile, "!!!! Self-referential !!!!"
            //     dodgy = any_err = 1
            if not dodgy:
                if bv >= 80:
                    shx1, shx2 = get_externsheet_local_range(bk, refx, blah)
                elif origrefx > 0:
                    shx1, shx2 = (-4, -4) // external ref
                else:
                    exty = bk._externsheet_type_b57[refx]
                    if exty == 4: // non-specific sheet in own doc't
                        shx1, shx2 = (-1, -1) // internal, any sheet
                    else:
                        shx1, shx2 = (-666, -666)
            okind = oUNK
            ovalue = None
            if shx1 == -5: // addin func name
                okind = oSTRG
                ovalue = bk.addin_func_names[tgtnamex]
                otext = '"' + ovalue.replace('"', '""') + '"'
            elif dodgy or shx1 < -1:
                otext = "<<Name //%d in external(?) file //%d>>" \
                        % (tgtnamex, origrefx)
            else:
                tgtobj = bk.name_obj_list[tgtnamex]
                if tgtobj.scope == -1:
                    otext = tgtobj.name
                else:
                    otext = "%s!%s" \
                            % (bk._sheet_names[tgtobj.scope], tgtobj.name)
                if blah:
                    print("    tNameX: setting text to", repr(res.text), file=bk.logfile)
            res = Operand(okind, ovalue, LEAF_RANK, otext)
            spush(res)
        elif opcode in error_opcodes:
            any_err = 1
            spush(error_opnd)
        else:
            if blah:
                print("FORMULA: /// Not handled yet: t" + oname, file=bk.logfile)
            any_err = 1
        if sz <= 0:
            raise FormulaError("Fatal: token size is not positive")
        pos += sz
    any_rel = not not any_rel
    if blah:
        print("End of formula. level=%d any_rel=%d any_err=%d stack=%r" % \
            (level, not not any_rel, any_err, stack), file=bk.logfile)
        if len(stack) >= 2:
            print("*** Stack has unprocessed args", file=bk.logfile)
        print(file=bk.logfile)

    if len(stack) != 1:
        result = None
    else:
        result = stack[0].text
    return result

//////// under deconstruction //////
def dump_formula(bk, data, fmlalen, bv, reldelta, blah=0, isname=0):
    if blah:
        print("dump_formula", fmlalen, bv, len(data), file=bk.logfile)
        hex_char_dump(data, 0, fmlalen, fout=bk.logfile)
    assert bv >= 80 //////// this function needs updating ////////
    sztab = szdict[bv]
    pos = 0
    stack = []
    any_rel = 0
    any_err = 0
    spush = stack.append
    while 0 <= pos < fmlalen:
        op = BYTES_ORD(data[pos])
        opcode = op & 0x1f
        optype = (op & 0x60) >> 5
        if optype:
            opx = opcode + 32
        else:
            opx = opcode
        oname = onames[opx] // + [" RVA"][optype]

        sz = sztab[opx]
        if blah:
            print("Pos:%d Op:0x%02x Name:t%s Sz:%d opcode:%02xh optype:%02xh" \
                % (pos, op, oname, sz, opcode, optype), file=bk.logfile)
        if not optype:
            if 0x01 <= opcode <= 0x02: // tExp, tTbl
                // reference to a shared formula or table record
                rowx, colx = unpack("<HH", data[pos+1:pos+5])
                if blah: print("  ", (rowx, colx), file=bk.logfile)
            elif opcode == 0x10: // tList
                if blah: print("tList pre", stack, file=bk.logfile)
                assert len(stack) >= 2
                bop = stack.pop()
                aop = stack.pop()
                spush(aop + bop)
                if blah: print("tlist post", stack, file=bk.logfile)
            elif opcode == 0x11: // tRange
                if blah: print("tRange pre", stack, file=bk.logfile)
                assert len(stack) >= 2
                bop = stack.pop()
                aop = stack.pop()
                assert len(aop) == 1
                assert len(bop) == 1
                result = do_box_funcs(tRangeFuncs, aop[0], bop[0])
                spush(result)
                if blah: print("tRange post", stack, file=bk.logfile)
            elif opcode == 0x0F: // tIsect
                if blah: print("tIsect pre", stack, file=bk.logfile)
                assert len(stack) >= 2
                bop = stack.pop()
                aop = stack.pop()
                assert len(aop) == 1
                assert len(bop) == 1
                result = do_box_funcs(tIsectFuncs, aop[0], bop[0])
                spush(result)
                if blah: print("tIsect post", stack, file=bk.logfile)
            elif opcode == 0x19: // tAttr
                subop, nc = unpack("<BH", data[pos+1:pos+4])
                subname = tAttrNames.get(subop, "??Unknown??")
                if subop == 0x04: // Choose
                    sz = nc * 2 + 6
                else:
                    sz = 4
                if blah: print("   subop=%02xh subname=t%s sz=%d nc=%02xh" % (subop, subname, sz, nc), file=bk.logfile)
            elif opcode == 0x17: // tStr
                if bv <= 70:
                    nc = BYTES_ORD(data[pos+1])
                    strg = data[pos+2:pos+2+nc] // left in 8-bit encoding
                    sz = nc + 2
                else:
                    strg, newpos = unpack_unicode_update_pos(data, pos+1, lenlen=1)
                    sz = newpos - pos
                if blah: print("   sz=%d strg=%r" % (sz, strg), file=bk.logfile)
            else:
                if sz <= 0:
                    print("**** Dud size; exiting ****", file=bk.logfile)
                    return
            pos += sz
            continue
        if opcode == 0x00: // tArray
            pass
        elif opcode == 0x01: // tFunc
            nb = 1 + int(bv >= 40)
            funcx = unpack("<" + " BH"[nb], data[pos+1:pos+1+nb])
            if blah: print("   FuncID=%d" % funcx, file=bk.logfile)
        elif opcode == 0x02: //tFuncVar
            nb = 1 + int(bv >= 40)
            nargs, funcx = unpack("<B" + " BH"[nb], data[pos+1:pos+2+nb])
            prompt, nargs = divmod(nargs, 128)
            macro, funcx = divmod(funcx, 32768)
            if blah: print("   FuncID=%d nargs=%d macro=%d prompt=%d" % (funcx, nargs, macro, prompt), file=bk.logfile)
        elif opcode == 0x03: //tName
            namex = unpack("<H", data[pos+1:pos+3])
            // Only change with BIFF version is the number of trailing UNUSED bytes!!!
            if blah: print("   namex=%d" % namex, file=bk.logfile)
        elif opcode == 0x04: // tRef
            res = get_cell_addr(data, pos+1, bv, reldelta)
            if blah: print("  ", res, file=bk.logfile)
        elif opcode == 0x05: // tArea
            res = get_cell_range_addr(data, pos+1, bv, reldelta)
            if blah: print("  ", res, file=bk.logfile)
        elif opcode == 0x09: // tMemFunc
            nb = unpack("<H", data[pos+1:pos+3])[0]
            if blah: print("  %d bytes of cell ref formula" % nb, file=bk.logfile)
        elif opcode == 0x0C: //tRefN
            res = get_cell_addr(data, pos+1, bv, reldelta=1)
            // note *ALL* tRefN usage has signed offset for relative addresses
            any_rel = 1
            if blah: print("   ", res, file=bk.logfile)
        elif opcode == 0x0D: //tAreaN
            res = get_cell_range_addr(data, pos+1, bv, reldelta=1)
            // note *ALL* tAreaN usage has signed offset for relative addresses
            any_rel = 1
            if blah: print("   ", res, file=bk.logfile)
        elif opcode == 0x1A: // tRef3d
            refx = unpack("<H", data[pos+1:pos+3])[0]
            res = get_cell_addr(data, pos+3, bv, reldelta)
            if blah: print("  ", refx, res, file=bk.logfile)
            rowx, colx, row_rel, col_rel = res
            any_rel = any_rel or row_rel or col_rel
            shx1, shx2 = get_externsheet_local_range(bk, refx, blah)
            any_err |= shx1 < -1
            coords = (shx1, shx2+1, rowx, rowx+1, colx, colx+1)
            if blah: print("   ", coords, file=bk.logfile)
            if optype == 1: spush([coords])
        elif opcode == 0x1B: // tArea3d
            refx = unpack("<H", data[pos+1:pos+3])[0]
            res1, res2 = get_cell_range_addr(data, pos+3, bv, reldelta)
            if blah: print("  ", refx, res1, res2, file=bk.logfile)
            rowx1, colx1, row_rel1, col_rel1 = res1
            rowx2, colx2, row_rel2, col_rel2 = res2
            any_rel = any_rel or row_rel1 or col_rel1 or row_rel2 or col_rel2
            shx1, shx2 = get_externsheet_local_range(bk, refx, blah)
            any_err |= shx1 < -1
            coords = (shx1, shx2+1, rowx1, rowx2+1, colx1, colx2+1)
            if blah: print("   ", coords, file=bk.logfile)
            if optype == 1: spush([coords])
        elif opcode == 0x19: // tNameX
            refx, namex = unpack("<HH", data[pos+1:pos+5])
            if blah: print("   refx=%d namex=%d" % (refx, namex), file=bk.logfile)
        elif opcode in error_opcodes:
            any_err = 1
        else:
            if blah: print("FORMULA: /// Not handled yet: t" + oname, file=bk.logfile)
            any_err = 1
        if sz <= 0:
            print("**** Dud size; exiting ****", file=bk.logfile)
            return
        pos += sz
    if blah:
        print("End of formula. any_rel=%d any_err=%d stack=%r" % \
            (not not any_rel, any_err, stack), file=bk.logfile)
        if len(stack) >= 2:
            print("*** Stack has unprocessed args", file=bk.logfile)

// === Some helper functions for displaying cell references ===

// I'm aware of only one possibility of a sheet-relative component in
// a reference: a 2D reference located in the "current sheet".
// xlrd stores this internally with bounds of (0, 1, ...) and
// relative flags of (1, 1, ...). These functions display the
// sheet component as empty, just like Excel etc.

def rownamerel(rowx, rowxrel, browx=None, r1c1=0):
    // if no base rowx is provided, we have to return r1c1
    if browx is None:
        r1c1 = True
    if not rowxrel:
        if r1c1:
            return "R%d" % (rowx+1)
        return "$%d" % (rowx+1)
    if r1c1:
        if rowx:
            return "R[%d]" % rowx
        return "R"
    return "%d" % ((browx + rowx) % 65536 + 1)

def colnamerel(colx, colxrel, bcolx=None, r1c1=0):
    // if no base colx is provided, we have to return r1c1
    if bcolx is None:
        r1c1 = True
    if not colxrel:
        if r1c1:
            return "C%d" % (colx + 1)
        return "$" + colname(colx)
    if r1c1:
        if colx:
            return "C[%d]" % colx
        return "C"
    return colname((bcolx + colx) % 256)

////
// Utility function: (5, 7) => 'H6'
def cellname(rowx, colx):
    """ (5, 7) => 'H6' """
    return "%s%d" % (colname(colx), rowx+1)

////
// Utility function: (5, 7) => '$H$6'
def cellnameabs(rowx, colx, r1c1=0):
    """ (5, 7) => '$H$6' or 'R8C6'"""
    if r1c1:
        return "R%dC%d" % (rowx+1, colx+1)
    return "$%s$%d" % (colname(colx), rowx+1)

def cellnamerel(rowx, colx, rowxrel, colxrel, browx=None, bcolx=None, r1c1=0):
    if not rowxrel and not colxrel:
        return cellnameabs(rowx, colx, r1c1)
    if (rowxrel and browx is None) or (colxrel and bcolx is None):
        // must flip the whole cell into R1C1 mode
        r1c1 = True
    c = colnamerel(colx, colxrel, bcolx, r1c1)
    r = rownamerel(rowx, rowxrel, browx, r1c1)
    if r1c1:
        return r + c
    return c + r

////
// Utility function: 7 => 'H', 27 => 'AB'
def colname(colx):
    """ 7 => 'H', 27 => 'AB' """
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if colx <= 25:
        return alphabet[colx]
    else:
        xdiv26, xmod26 = divmod(colx, 26)
        return alphabet[xdiv26 - 1] + alphabet[xmod26]

def rangename2d(rlo, rhi, clo, chi, r1c1=0):
    """ (5, 20, 7, 10) => '$H$6:$J$20' """
    if r1c1:
        return
    if rhi == rlo+1 and chi == clo+1:
        return cellnameabs(rlo, clo, r1c1)
    return "%s:%s" % (cellnameabs(rlo, clo, r1c1), cellnameabs(rhi-1, chi-1, r1c1))

def rangename2drel(rlo_rhi_clo_chi, rlorel_rhirel_clorel_chirel, browx=None, bcolx=None, r1c1=0):
    rlo, rhi, clo, chi = rlo_rhi_clo_chi
    rlorel, rhirel, clorel, chirel = rlorel_rhirel_clorel_chirel
    if (rlorel or rhirel) and browx is None:
        r1c1 = True
    if (clorel or chirel) and bcolx is None:
        r1c1 = True
    return "%s:%s" % (
        cellnamerel(rlo,   clo,   rlorel, clorel, browx, bcolx, r1c1),
        cellnamerel(rhi-1, chi-1, rhirel, chirel, browx, bcolx, r1c1)
        )
////
// Utility function:
// <br /> Ref3D((1, 4, 5, 20, 7, 10)) => 'Sheet2:Sheet3!$H$6:$J$20'
def rangename3d(book, ref3d):
    """ Ref3D(1, 4, 5, 20, 7, 10) => 'Sheet2:Sheet3!$H$6:$J$20'
        (assuming Excel's default sheetnames) """
    coords = ref3d.coords
    return "%s!%s" % (
        sheetrange(book, *coords[:2]),
        rangename2d(*coords[2:6]))

////
// Utility function:
// <br /> Ref3D(coords=(0, 1, -32, -22, -13, 13), relflags=(0, 0, 1, 1, 1, 1))
// R1C1 mode => 'Sheet1!R[-32]C[-13]:R[-23]C[12]'
// A1 mode => depends on base cell (browx, bcolx)
def rangename3drel(book, ref3d, browx=None, bcolx=None, r1c1=0):
    coords = ref3d.coords
    relflags = ref3d.relflags
    shdesc = sheetrangerel(book, coords[:2], relflags[:2])
    rngdesc = rangename2drel(coords[2:6], relflags[2:6], browx, bcolx, r1c1)
    if not shdesc:
        return rngdesc
    return "%s!%s" % (shdesc, rngdesc)
*/

static std::map<int, std::string>
shname_dict_ = {
    {-1, "?internal; any sheet?"},
    {-2, "internal; deleted sheet"},
    {-3, "internal; macro sheet"},
    {-4, "<<external>>"},
};

inline
std::string
quotedsheetname(std::vector<std::string> shnames, int shx)
{
    std::string shname;
    if (shx >= 0) {
        shname = shnames[shx];
    }
    else {
        shname = utils::getelse(shname_dict_, shx, strutil::format("?error %d?", shx));
    }
    if (shname.find("'") != std::string::npos) {
        return strutil::format("'%s'", strutil::replace(shname, "'", "''"));
    }
    if (shname.find(" ") != std::string::npos) {
        return strutil::format("'%s'", shname);
    }
    return shname;
}


class FormulaSheetNamesDelegate
{
public:
    std::vector<std::string> sheet_names();
};

inline
std::string sheetrange(FormulaSheetNamesDelegate& book, int slo, int shi)
{
    auto shnames = book.sheet_names();
    auto shdesc = quotedsheetname(shnames, slo);
    if (slo != shi-1) {
        shdesc += ":" + quotedsheetname(shnames, shi-1);
    }
    return shdesc;
}

inline
std::string sheetrangerel(FormulaSheetNamesDelegate& book,
                          std::tuple<int, int> srange,
                          std::tuple<int, int> srangerel)
{
    int slo, shi, slorel, shirel;
    std::tie(slo, shi) = srange;
    std::tie(slorel, shirel) = srangerel;
    if (!slorel && !shirel) {
        return sheetrange(book, slo, shi);
    }
    // assert (slo == 0 == shi-1) and slorel and shirel
    return "";
}

// ==============================================================

}
}