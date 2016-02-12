#pragma once
// -*- coding: cp1252 -*-

////
// Support module for the xlrd package.
//
// <p>Portions copyright © 2005-2010 Stephen John Machin, Lingfo Pty Ltd</p>
// <p>This module is part of the xlrd package, which is released under a BSD-style licence.</p>
////

// 2010-03-01 SJM Reading SCL record
// 2010-03-01 SJM Added more record IDs for biff_dump & biff_count
// 2008-02-10 SJM BIFF2 BLANK record
// 2008-02-08 SJM Preparation for Excel 2.0 support
// 2008-02-02 SJM Added suffixes (_B2, _B2_ONLY, etc) on record names for biff_dump & biff_count
// 2007-12-04 SJM Added support for Excel 2.x (BIFF2) files.
// 2007-09-08 SJM Avoid crash when zero-length Unicode string missing options byte.
// 2007-04-22 SJM Remove experimental "trimming" facility.

#include <string>
#include <vector>
#include <map>
#include <exception>

#include "./utils.h"

namespace xlrd {
namespace biffh {

namespace strutil = utils::str;

static int DEBUG = 0;


class XLRDError : public std::runtime_error
{
};

enum {
    FUN, FDT, FNU, FGE, FTX
    // unknown, date, number, general, text
};
static const int DATEFORMAT = FDT;
static const int NUMBERFORMAT = FNU;

enum {
    XL_CELL_EMPTY,
    XL_CELL_TEXT,
    XL_CELL_NUMBER,
    XL_CELL_DATE,
    XL_CELL_BOOLEAN,
    XL_CELL_ERROR,
    XL_CELL_BLANK  // for use in debugging, gathering stats, etc
};

std::map<int, std::string>
biff_text_from_num = {
    {0,  "(not BIFF)"},
    {20, "2.0"},
    {21, "2.1"},
    {30, "3"},
    {40, "4S"},
    {45, "4W"},
    {50, "5"},
    {70, "7"},
    {80, "8"},
    {85, "8X"}
};

////
// <p>This dictionary can be used to produce a text version of the internal codes
// that Excel uses for error cells. Here are its contents:
// <pre>
// 0x00: '//NULL!',  // Intersection of two cell ranges is empty
// 0x07: '//DIV/0!', // Division by zero
// 0x0F: '//VALUE!', // Wrong type of operand
// 0x17: '//REF!',   // Illegal or deleted cell reference
// 0x1D: '//NAME?',  // Wrong function or range name
// 0x24: '//NUM!',   // Value range overflow
// 0x2A: '//N/A',    // Argument or function not available
// </pre></p>

std::map<int, std::string>
error_text_from_code = {
    {0x00, "//NULL!"},   // Intersection of two cell ranges is empty
    {0x07, "//DIV/0!"},  // Division by zero
    {0x0F, "//VALUE!"},  // Wrong type of operand
    {0x17, "//REF!"},    // Illegal or deleted cell reference
    {0x1D, "//NAME?"},   // Wrong function or range name
    {0x24, "//NUM!"},    // Value range overflow
    {0x2A, "//N/A"}      // Argument or function not available
};

static int BIFF_FIRST_UNICODE = 80;

static int WBKBLOBAL = 0x5;
static int XL_WORKBOOK_GLOBALS = WBKBLOBAL;
static int XL_WORKBOOK_GLOBALS_4W = 0x100;
static int WRKSHEET = 0x10;
static int XL_WORKSHEET = WRKSHEET;

static int XL_BOUNDSHEET_WORKSHEET = 0x00;
static int XL_BOUNDSHEET_CHART     = 0x02;
static int XL_BOUNDSHEET_VB_MODULE = 0x06;

// XL_RK2 = 0x7e
static int XL_ARRAY  = 0x0221;
static int XL_ARRAY2 = 0x0021;
static int XL_BLANK = 0x0201;
static int XL_BLANK_B2 = 0x01;
static int XL_BOF = 0x809;
static int XL_BOOLERR = 0x205;
static int XL_BOOLERR_B2 = 0x5;
static int XL_BOUNDSHEET = 0x85;
static int XL_BUILTINFMTCOUNT = 0x56;
static int XL_CF = 0x01B1;
static int XL_CODEPAGE = 0x42;
static int XL_COLINFO = 0x7D;
static int XL_COLUMNDEFAULT = 0x20; // BIFF2 only
static int XL_COLWIDTH = 0x24; // BIFF2 only
static int XL_CONDFMT = 0x01B0;
static int XL_CONTINUE = 0x3c;
static int XL_COUNTRY = 0x8C;
static int XL_DATEMODE = 0x22;
static int XL_DEFAULTROWHEIGHT = 0x0225;
static int XL_DEFCOLWIDTH = 0x55;
static int XL_DIMENSION = 0x200;
static int XL_DIMENSION2 = 0x0;
static int XL_EFONT = 0x45;
static int XL_EOF = 0x0a;
static int XL_EXTERNNAME = 0x23;
static int XL_EXTERNSHEET = 0x17;
static int XL_EXTSST = 0xff;
static int XL_FEAT11 = 0x872;
static int XL_FILEPASS = 0x2f;
static int XL_FONT = 0x31;
static int XL_FONT_B3B4 = 0x231;
static int XL_FORMAT = 0x41e;
static int XL_FORMAT2 = 0x1E; // BIFF2, BIFF3
static int XL_FORMULA = 0x6;
static int XL_FORMULA3 = 0x206;
static int XL_FORMULA4 = 0x406;
static int XL_GCW = 0xab;
static int XL_HLINK = 0x01B8;
static int XL_QUICKTIP = 0x0800;
static int XL_HORIZONTALPAGEBREAKS = 0x1b;
static int XL_INDEX = 0x20b;
static int XL_INTEGER = 0x2; // BIFF2 only
static int XL_IXFE = 0x44; // BIFF2 only
static int XL_LABEL = 0x204;
static int XL_LABEL_B2 = 0x04;
static int XL_LABELRANGES = 0x15f;
static int XL_LABELSST = 0xfd;
static int XL_LEFTMARGIN = 0x26;
static int XL_TOPMARGIN = 0x28;
static int XL_RIGHTMARGIN = 0x27;
static int XL_BOTTOMMARGIN = 0x29;
static int XL_HEADER = 0x14;
static int XL_FOOTER = 0x15;
static int XL_HCENTER = 0x83;
static int XL_VCENTER = 0x84;
static int XL_MERGEDCELLS = 0xE5;
static int XL_MSO_DRAWING = 0x00EC;
static int XL_MSO_DRAWING_GROUP = 0x00EB;
static int XL_MSO_DRAWING_SELECTION = 0x00ED;
static int XL_MULRK = 0xbd;
static int XL_MULBLANK = 0xbe;
static int XL_NAME = 0x18;
static int XL_NOTE = 0x1c;
static int XL_NUMBER = 0x203;
static int XL_NUMBER_B2 = 0x3;
static int XL_OBJ = 0x5D;
static int XL_PAGESETUP = 0xA1;
static int XL_PALETTE = 0x92;
static int XL_PANE = 0x41;
static int XL_PRINTGRIDLINES = 0x2B;
static int XL_PRINTHEADERS = 0x2A;
static int XL_RK = 0x27e;
static int XL_ROW = 0x208;
static int XL_ROW_B2 = 0x08;
static int XL_RSTRING = 0xd6;
static int XL_SCL = 0x00A0;
static int XL_SHEETHDR = 0x8F; // BIFF4W only
static int XL_SHEETPR = 0x81;
static int XL_SHEETSOFFSET = 0x8E; // BIFF4W only
static int XL_SHRFMLA = 0x04bc;
static int XL_SST = 0xfc;
static int XL_STANDARDWIDTH = 0x99;
static int XL_STRING = 0x207;
static int XL_STRING_B2 = 0x7;
static int XL_STYLE = 0x293;
static int XL_SUPBOOK = 0x1AE; // aka EXTERNALBOOK in OOo docs
static int XL_TABLEOP = 0x236;
static int XL_TABLEOP2 = 0x37;
static int XL_TABLEOP_B2 = 0x36;
static int XL_TXO = 0x1b6;
static int XL_UNCALCED = 0x5e;
static int XL_UNKNOWN = 0xffff;
static int XL_VERTICALPAGEBREAKS = 0x1a;
static int XL_WINDOW2    = 0x023E;
static int XL_WINDOW2_B2 = 0x003E;
static int XL_WRITEACCESS = 0x5C;
static int XL_WSBOOL = XL_SHEETPR;
static int XL_XF = 0xe0;
static int XL_XF2 = 0x0043; // BIFF2 version of XF record
static int XL_XF3 = 0x0243; // BIFF3 version of XF record
static int XL_XF4 = 0x0443; // BIFF4 version of XF record

static std::map<int, int>
boflen = {{0x0809, 8}, {0x0409, 6}, {0x0209, 6}, {0x0009, 4}};
static std::vector<int>
bofcodes = {0x0809, 0x0409, 0x0209, 0x0009};

static std::vector<int>
XL_FORMULA_OPCODES = {0x0006, 0x0406, 0x0206};

static std::vector<int>
_cell_opcode_list = {
    XL_BOOLERR,
    XL_FORMULA,
    XL_FORMULA3,
    XL_FORMULA4,
    XL_LABEL,
    XL_LABELSST,
    XL_MULRK,
    XL_NUMBER,
    XL_RK,
    XL_RSTRING,
};
static std::map<int, int>
_cell_opcode_dict;
//for _cell_opcode in _cell_opcode_list:
//    _cell_opcode_dict[_cell_opcode] = 1

inline
bool
is_cell_opcode(int c) {
    if (_cell_opcode_dict.empty()) {
        for (auto _cell_opcode: _cell_opcode_list) {
            _cell_opcode_dict[_cell_opcode] = 1;
        }
    }
    return _cell_opcode_dict.find(c) != _cell_opcode_dict.end();
}

/*
def upkbits(tgt_obj, src, manifest, local_setattr=setattr):
    for n, mask, attr in manifest:
        local_setattr(tgt_obj, attr, (src & mask) >> n)

def upkbitsL(tgt_obj, src, manifest, local_setattr=setattr, local_int=int):
    for n, mask, attr in manifest:
        local_setattr(tgt_obj, attr, local_int((src & mask) >> n))

def unpack_string(data, pos, encoding, lenlen=1):
    nchars = unpack('<' + 'BH'[lenlen-1], data[pos:pos+lenlen])[0]
    pos += lenlen
    return unicode(data[pos:pos+nchars], encoding)
*/

inline
std::tuple<std::string, int>
unpack_string_update_pos(std::vector<uint8_t> data, int pos,
                         std::string encoding, int lenlen=1,
                         int known_len=-1)
{
    int nchars = 0;
    if (known_len > -1) {
        // On a NAME record, the length byte is detached from the front of the string.
        nchars = known_len;
    }
    else {
        if (lenlen == 1) {
            nchars = utils::as_uint8(data, pos);
        } else {
            nchars = utils::as_uint16(data, pos);
        }
        pos += lenlen;
    }
    int newpos = pos + nchars;
    return std::make_tuple(strutil::unicode(utils::slice(data, pos, newpos), encoding), newpos);
}

/*
def unpack_unicode(data, pos, lenlen=2):
    "Return unicode_strg"
    nchars = unpack('<' + 'BH'[lenlen-1], data[pos:pos+lenlen])[0]
    if not nchars:
        // Ambiguous whether 0-length string should have an "options" byte.
        // Avoid crash if missing.
        return UNICODE_LITERAL("")
    pos += lenlen
    options = BYTES_ORD(data[pos])
    pos += 1
    // phonetic = options & 0x04
    // richtext = options & 0x08
    if options & 0x08:
        // rt = unpack('<H', data[pos:pos+2])[0] // unused
        pos += 2
    if options & 0x04:
        // sz = unpack('<i', data[pos:pos+4])[0] // unused
        pos += 4
    if options & 0x01:
        // Uncompressed UTF-16-LE
        rawstrg = data[pos:pos+2*nchars]
        // if DEBUG: print "nchars=%d pos=%d rawstrg=%r" % (nchars, pos, rawstrg)
        strg = unicode(rawstrg, 'utf_16_le')
        // pos += 2*nchars
    else:
        // Note: this is COMPRESSED (not ASCII!) encoding!!!
        // Merely returning the raw bytes would work OK 99.99% of the time
        // if the local codepage was cp1252 -- however this would rapidly go pear-shaped
        // for other codepages so we grit our Anglocentric teeth and return Unicode :-)

        strg = unicode(data[pos:pos+nchars], "latin_1")
        // pos += nchars
    // if richtext:
    //     pos += 4 * rt
    // if phonetic:
    //     pos += sz
    // return (strg, pos)
    return strg
*/

inline
std::tuple<std::string, int>
unpack_unicode_update_pos(std::vector<uint8_t> data, int pos, int lenlen=2, int known_len=-1) {
    // "Return (unicode_strg, updated value of pos)"
    int nchars;
    if (known_len > -1) {
        // On a NAME record, the length byte is detached from the front of the string.
        nchars = known_len;
    }
    else {
        //nchars = unpack('<' + 'BH'[lenlen-1], data[pos:pos+lenlen])[0]
        if (lenlen == 1) {
            nchars = utils::as_uint8(data, pos);
        } else {
            nchars = utils::as_uint16(data, pos);
        }
        pos += lenlen;
    }
    if (nchars == 0 && data.size() < pos) {
        // Zero-length string with no options byte
        return std::make_tuple("", pos);
    }
    uint8_t options = data[pos];
    pos += 1;
    int phonetic = options & 0x04;
    int richtext = options & 0x08;
    int rt = 0;
    if (richtext) {
        // rt = unpack('<H', data[pos:pos+2])[0]
        rt = utils::as_uint16(data, pos);
        pos += 2;
    }
    int sz = 0;
    if (phonetic) {
        // sz = unpack('<i', data[pos:pos+4])[0]
        sz = utils::as_int32(data, pos);
        pos += 4;
    }
    std::string strg;
    if (options & 0x01) {
        // Uncompressed UTF-16-LE
        // strg = unicode(data[pos:pos+2*nchars], 'utf_16_le')
        strg = strutil::utf16to8(utils::slice(data, pos, pos+2*nchars));
        pos += 2*nchars;
    }
    else {
        // Note: this is COMPRESSED (not ASCII!) encoding!!!
        // strg = unicode(data[pos:pos+nchars], "latin_1")
        strg = std::string((char*)&utils::slice(data, pos, pos+nchars)[0]);
        pos += nchars;
    }
    if (richtext) {
        pos += 4 * rt;
    }
    if (phonetic) {
        pos += sz;
    }
    return std::make_tuple(strg, pos);
}

inline
int
unpack_cell_range_address_list_update_pos(
    std::vector<std::tuple<int, int, int, int>>* output_list,
    std::vector<uint8_t>& data,
    int pos, int biff_version,
    int addr_size=6
) {
    // output_list is updated in situ
    // assert addr_size in (6, 8)
    // Used to assert size == 6 if not BIFF8, but pyWLWriter writes
    // BIFF8-only MERGEDCELLS records in a BIFF5 file!
    int n = utils::as_uint16(data, pos);
    pos += 2;
    if (n) {
        for (int i=0; i < n; i++) {
            int ra, rb, ca, cb;
            if (addr_size == 6) {
                ra = utils::as_uint16(data, pos);
                rb = utils::as_uint16(data, pos+2);
                ca = utils::as_uint8(data, pos+4);
                cb = utils::as_uint8(data, pos+5);
            } else {  // addr_size == 8
                ra = utils::as_uint16(data, pos);
                rb = utils::as_uint16(data, pos+2);
                ca = utils::as_uint16(data, pos+4);
                cb = utils::as_uint16(data, pos+6);
            }
            output_list->push_back(std::make_tuple(ra, rb+1, ca, cb+1));
            pos += addr_size;
        }
    }
    return pos;
}

/*
*/

static
std::map<int, std::string>
biff_rec_name_dict = {
    {0x0000, "DIMENSIONS_B2"},
    {0x0001, "BLANK_B2"},
    {0x0002, "INTEGER_B2_ONLY"},
    {0x0003, "NUMBER_B2"},
    {0x0004, "LABEL_B2"},
    {0x0005, "BOOLERR_B2"},
    {0x0006, "FORMULA"},
    {0x0007, "STRING_B2"},
    {0x0008, "ROW_B2"},
    {0x0009, "BOF_B2"},
    {0x000A, "EOF"},
    {0x000B, "INDEX_B2_ONLY"},
    {0x000C, "CALCCOUNT"},
    {0x000D, "CALCMODE"},
    {0x000E, "PRECISION"},
    {0x000F, "REFMODE"},
    {0x0010, "DELTA"},
    {0x0011, "ITERATION"},
    {0x0012, "PROTECT"},
    {0x0013, "PASSWORD"},
    {0x0014, "HEADER"},
    {0x0015, "FOOTER"},
    {0x0016, "EXTERNCOUNT"},
    {0x0017, "EXTERNSHEET"},
    {0x0018, "NAME_B2,5+"},
    {0x0019, "WINDOWPROTECT"},
    {0x001A, "VERTICALPAGEBREAKS"},
    {0x001B, "HORIZONTALPAGEBREAKS"},
    {0x001C, "NOTE"},
    {0x001D, "SELECTION"},
    {0x001E, "FORMAT_B2-3"},
    {0x001F, "BUILTINFMTCOUNT_B2"},
    {0x0020, "COLUMNDEFAULT_B2_ONLY"},
    {0x0021, "ARRAY_B2_ONLY"},
    {0x0022, "DATEMODE"},
    {0x0023, "EXTERNNAME"},
    {0x0024, "COLWIDTH_B2_ONLY"},
    {0x0025, "DEFAULTROWHEIGHT_B2_ONLY"},
    {0x0026, "LEFTMARGIN"},
    {0x0027, "RIGHTMARGIN"},
    {0x0028, "TOPMARGIN"},
    {0x0029, "BOTTOMMARGIN"},
    {0x002A, "PRINTHEADERS"},
    {0x002B, "PRINTGRIDLINES"},
    {0x002F, "FILEPASS"},
    {0x0031, "FONT"},
    {0x0032, "FONT2_B2_ONLY"},
    {0x0036, "TABLEOP_B2"},
    {0x0037, "TABLEOP2_B2"},
    {0x003C, "CONTINUE"},
    {0x003D, "WINDOW1"},
    {0x003E, "WINDOW2_B2"},
    {0x0040, "BACKUP"},
    {0x0041, "PANE"},
    {0x0042, "CODEPAGE"},
    {0x0043, "XF_B2"},
    {0x0044, "IXFE_B2_ONLY"},
    {0x0045, "EFONT_B2_ONLY"},
    {0x004D, "PLS"},
    {0x0051, "DCONREF"},
    {0x0055, "DEFCOLWIDTH"},
    {0x0056, "BUILTINFMTCOUNT_B3-4"},
    {0x0059, "XCT"},
    {0x005A, "CRN"},
    {0x005B, "FILESHARING"},
    {0x005C, "WRITEACCESS"},
    {0x005D, "OBJECT"},
    {0x005E, "UNCALCED"},
    {0x005F, "SAVERECALC"},
    {0x0063, "OBJECTPROTECT"},
    {0x007D, "COLINFO"},
    {0x007E, "RK2_mythical_?"},
    {0x0080, "GUTS"},
    {0x0081, "WSBOOL"},
    {0x0082, "GRIDSET"},
    {0x0083, "HCENTER"},
    {0x0084, "VCENTER"},
    {0x0085, "BOUNDSHEET"},
    {0x0086, "WRITEPROT"},
    {0x008C, "COUNTRY"},
    {0x008D, "HIDEOBJ"},
    {0x008E, "SHEETSOFFSET"},
    {0x008F, "SHEETHDR"},
    {0x0090, "SORT"},
    {0x0092, "PALETTE"},
    {0x0099, "STANDARDWIDTH"},
    {0x009B, "FILTERMODE"},
    {0x009C, "FNGROUPCOUNT"},
    {0x009D, "AUTOFILTERINFO"},
    {0x009E, "AUTOFILTER"},
    {0x00A0, "SCL"},
    {0x00A1, "SETUP"},
    {0x00AB, "GCW"},
    {0x00BD, "MULRK"},
    {0x00BE, "MULBLANK"},
    {0x00C1, "MMS"},
    {0x00D6, "RSTRING"},
    {0x00D7, "DBCELL"},
    {0x00DA, "BOOKBOOL"},
    {0x00DD, "SCENPROTECT"},
    {0x00E0, "XF"},
    {0x00E1, "INTERFACEHDR"},
    {0x00E2, "INTERFACEEND"},
    {0x00E5, "MERGEDCELLS"},
    {0x00E9, "BITMAP"},
    {0x00EB, "MSO_DRAWING_GROUP"},
    {0x00EC, "MSO_DRAWING"},
    {0x00ED, "MSO_DRAWING_SELECTION"},
    {0x00EF, "PHONETIC"},
    {0x00FC, "SST"},
    {0x00FD, "LABELSST"},
    {0x00FF, "EXTSST"},
    {0x013D, "TABID"},
    {0x015F, "LABELRANGES"},
    {0x0160, "USESELFS"},
    {0x0161, "DSF"},
    {0x01AE, "SUPBOOK"},
    {0x01AF, "PROTECTIONREV4"},
    {0x01B0, "CONDFMT"},
    {0x01B1, "CF"},
    {0x01B2, "DVAL"},
    {0x01B6, "TXO"},
    {0x01B7, "REFRESHALL"},
    {0x01B8, "HLINK"},
    {0x01BC, "PASSWORDREV4"},
    {0x01BE, "DV"},
    {0x01C0, "XL9FILE"},
    {0x01C1, "RECALCID"},
    {0x0200, "DIMENSIONS"},
    {0x0201, "BLANK"},
    {0x0203, "NUMBER"},
    {0x0204, "LABEL"},
    {0x0205, "BOOLERR"},
    {0x0206, "FORMULA_B3"},
    {0x0207, "STRING"},
    {0x0208, "ROW"},
    {0x0209, "BOF"},
    {0x020B, "INDEX_B3+"},
    {0x0218, "NAME"},
    {0x0221, "ARRAY"},
    {0x0223, "EXTERNNAME_B3-4"},
    {0x0225, "DEFAULTROWHEIGHT"},
    {0x0231, "FONT_B3B4"},
    {0x0236, "TABLEOP"},
    {0x023E, "WINDOW2"},
    {0x0243, "XF_B3"},
    {0x027E, "RK"},
    {0x0293, "STYLE"},
    {0x0406, "FORMULA_B4"},
    {0x0409, "BOF"},
    {0x041E, "FORMAT"},
    {0x0443, "XF_B4"},
    {0x04BC, "SHRFMLA"},
    {0x0800, "QUICKTIP"},
    {0x0809, "BOF"},
    {0x0862, "SHEETLAYOUT"},
    {0x0867, "SHEETPROTECTION"},
    {0x0868, "RANGEPROTECTION"},
};

/*
def hex_char_dump(strg, ofs, dlen, base=0, fout=sys.stdout, unnumbered=False):
    endpos = min(ofs + dlen, len(strg))
    pos = ofs
    numbered = not unnumbered
    num_prefix = ''
    while pos < endpos:
        endsub = min(pos + 16, endpos)
        substrg = strg[pos:endsub]
        lensub = endsub - pos
        if lensub <= 0 or lensub != len(substrg):
            fprintf(
                sys.stdout,
                '??? hex_char_dump: ofs=%d dlen=%d base=%d -> endpos=%d pos=%d endsub=%d substrg=%r\n',
                ofs, dlen, base, endpos, pos, endsub, substrg)
            break
        hexd = ''.join(["%02x " % BYTES_ORD(c) for c in substrg])
        
        chard = ''
        for c in substrg:
            c = chr(BYTES_ORD(c))
            if c == '\0':
                c = '~'
            elif not (' ' <= c <= '~'):
                c = '?'
            chard += c
        if numbered:
            num_prefix = "%5d: " %  (base+pos-ofs)
        
        fprintf(fout, "%s     %-48s %s\n", num_prefix, hexd, chard)
        pos = endsub

def biff_dump(mem, stream_offset, stream_len, base=0, fout=sys.stdout, unnumbered=False):
    pos = stream_offset
    stream_end = stream_offset + stream_len
    adj = base - stream_offset
    dummies = 0
    numbered = not unnumbered
    num_prefix = ''
    while stream_end - pos >= 4:
        rc, length = unpack('<HH', mem[pos:pos+4])
        if rc == 0 and length == 0:
            if mem[pos:] == b'\0' * (stream_end - pos):
                dummies = stream_end - pos
                savpos = pos
                pos = stream_end
                break
            if dummies:
                dummies += 4
            else:
                savpos = pos
                dummies = 4
            pos += 4
        else:
            if dummies:
                if numbered:
                    num_prefix =  "%5d: " % (adj + savpos)
                fprintf(fout, "%s---- %d zero bytes skipped ----\n", num_prefix, dummies)
                dummies = 0
            recname = biff_rec_name_dict.get(rc, '<UNKNOWN>')
            if numbered:
                num_prefix = "%5d: " % (adj + pos)
            fprintf(fout, "%s%04x %s len = %04x (%d)\n", num_prefix, rc, recname, length, length)
            pos += 4
            hex_char_dump(mem, pos, length, adj+pos, fout, unnumbered)
            pos += length
    if dummies:
        if numbered:
            num_prefix =  "%5d: " % (adj + savpos)
        fprintf(fout, "%s---- %d zero bytes skipped ----\n", num_prefix, dummies)
    if pos < stream_end:
        if numbered:
            num_prefix = "%5d: " % (adj + pos)
        fprintf(fout, "%s---- Misc bytes at end ----\n", num_prefix)
        hex_char_dump(mem, pos, stream_end-pos, adj + pos, fout, unnumbered)
    elif pos > stream_end:
        fprintf(fout, "Last dumped record has length (%d) that is too large\n", length)

def biff_count_records(mem, stream_offset, stream_len, fout=sys.stdout):
    pos = stream_offset
    stream_end = stream_offset + stream_len
    tally = {}
    while stream_end - pos >= 4:
        rc, length = unpack('<HH', mem[pos:pos+4])
        if rc == 0 and length == 0:
            if mem[pos:] == b'\0' * (stream_end - pos):
                break
            recname = "<Dummy (zero)>"
        else:
            recname = biff_rec_name_dict.get(rc, None)
            if recname is None:
                recname = "Unknown_0x%04X" % rc
        if recname in tally:
            tally[recname] += 1
        else:
            tally[recname] = 1
        pos += length + 4
    slist = sorted(tally.items())
    for recname, count in slist:
        print("%8d %s" % (count, recname), file=fout)

encoding_from_codepage = {
    1200 : 'utf_16_le',
    10000: 'mac_roman',
    10006: 'mac_greek', // guess
    10007: 'mac_cyrillic', // guess
    10029: 'mac_latin2', // guess
    10079: 'mac_iceland', // guess
    10081: 'mac_turkish', // guess
    32768: 'mac_roman',
    32769: 'cp1252',
    }
// some more guessing, for Indic scripts
// codepage 57000 range:
// 2 Devanagari [0]
// 3 Bengali [1]
// 4 Tamil [5]
// 5 Telegu [6]
// 6 Assamese [1] c.f. Bengali
// 7 Oriya [4]
// 8 Kannada [7]
// 9 Malayalam [8]
// 10 Gujarati [3]
// 11 Gurmukhi [2]

*/

}
}
