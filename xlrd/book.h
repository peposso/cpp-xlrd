#pragma once
/*
# Copyright (c) 2005-2012 Stephen John Machin, Lingfo Pty Ltd
# This module is part of the xlrd package, which is released under a
# BSD-style licence.

from __future__ import print_function

from .timemachine import *
from .biffh import *
import struct; unpack = struct.unpack
import sys
import time
from . import sheet
from . import compdoc
from .formula import *
from . import formatting
if sys.version.startswith("IronPython"):
    # print >> sys.stderr, "...importing encodings"
    import encodings

empty_cell = sheet.empty_cell # for exposure to the world ...

*/

#include "./biffh.h"
#include "./sheet.h"
#include "./compdoc.h"
#include "./formula.h"  // __all__

#include <string>
#include <vector>
#include <map>
#include <tuple>

namespace xlrd {
namespace book {

namespace strutil = utils::str;
USING_UTILS_PP;

using Operand = formula::Operand;
using Sheet = sheet::Sheet;
using Cell = sheet::Cell;
using XLRDError = biffh::XLRDError;

int DEBUG = 0;

int USE_FANCY_CD = 1;

int TOGGLE_GC = 0;

// import gc
// gc.set_debug(gc.DEBUG_STATS)

// try:
//     import mmap
//     MMAP_AVAILABLE = 1
// except ImportError:
//     MMAP_AVAILABLE = 0
// USE_MMAP = MMAP_AVAILABLE

int MY_EOF = 0xF00BAAA;  // not a 16-bit number

enum{
    SUPBOOK_UNK, SUPBOOK_INTERNAL, SUPBOOK_EXTERNAL, SUPBOOK_ADDIN, SUPBOOK_DDEOLE
};

std::vector<int>
SUPPORTED_VERSIONS = {80, 70, 50, 45, 40, 30, 21, 20};

std::map<std::string, char>
_code_from_builtin_name = {
    {"Consolidate_Area", '\x00'},
    {"Auto_Open",        '\x01'},
    {"Auto_Close",       '\x02'},
    {"Extract",          '\x03'},
    {"Database",         '\x04'},
    {"Criteria",         '\x05'},
    {"Print_Area",       '\x06'},
    {"Print_Titles",     '\x07'},
    {"Recorder",         '\x08'},
    {"Data_Form",        '\x09'},
    {"Auto_Activate",    '\x0A'},
    {"Auto_Deactivate",  '\x0B'},
    {"Sheet_Title",      '\x0C'},
    {"_FilterDatabase",  '\x0D'},
};


class Book;

class Name: public formula::FormulaNameDelegate {
public:
    Book* book;

    ////
    // 0 = Visible; 1 = Hidden
    int hidden = 0;

    ////
    // 0 = Command macro; 1 = Function macro. Relevant only if macro == 1
    int func = 0;

    ////
    // 0 = Sheet macro; 1 = VisualBasic macro. Relevant only if macro == 1
    int vbasic = 0;

    ////
    // 0 = Standard name; 1 = Macro name
    int macro = 0;

    ////
    // 0 = Simple formula; 1 = Complex formula (array formula or user defined)<br />
    // <i>No examples have been sighted.</i>
    int complex = 0;

    ////
    // 0 = User-defined name; 1 = Built-in name
    // (common examples: Print_Area, Print_Titles; see OOo docs for full list)
    int builtin = 0;

    ////
    // Function group. Relevant only if macro == 1; see OOo docs for values.
    int funcgroup = 0;

    ////
    // 0 = Formula definition; 1 = Binary data<br />  <i>No examples have been sighted.</i>
    int binary = 0;

    ////
    // The index of this object in book.name_obj_list
    int name_index = 0;

    ////
    // A Unicode string. If builtin, decoded as per OOo docs.
    std::string name;  // = "";

    ////
    // An 8-bit string.
    std::vector<uint8_t> raw_formula;

    ////
    // -1: The name is global (visible in all calculation sheets).<br />
    // -2: The name belongs to a macro sheet or VBA sheet.<br />
    // -3: The name is invalid.<br />
    // 0 <= scope < book.nsheets: The name is local to the sheet whose index is scope.
    int scope = -1;

    ////
    // The result of evaluating the formula, if any.
    // If no formula, or evaluation of the formula encountered problems,
    // the result is None. Otherwise the result is a single instance of the
    // Operand class.
    //
    formula::Operand* result;  // = None

    ////
    // This is a convenience method for the frequent use case where the name
    // refers to a single cell.
    // @return An instance of the Cell class.
    // @throws XLRDError The name is not a constant absolute reference
    // to a single cell.
    inline
    sheet::Cell cell()
    {
        // Operand* res = this->result;
        // if (res != nullptr) {
        //     // result should be an instance of the Operand class
        //     int kind = res->kind;
        //     auto value = res->value;
        //     // if (kind == formula::oREF && value.size() == 1) {
        //     //     ref3d = value[0];
        //     //     if ((0 <= ref3d.shtxlo) &&
        //     //         (ref3d.shtxlo == ref3d.shtxhi - 1) &&
        //     //         (ref3d.rowxlo == ref3d.rowxhi - 1) &&
        //     //         (ref3d.colxlo == ref3d.colxhi - 1))
        //     //     {
        //     //         sh = self.book.sheet_by_index(ref3d.shtxlo);
        //     //         return sh.cell(ref3d.rowxlo, ref3d.colxlo);
        //     //     }
        //     // }
        // }
        // self.dump(self.book.logfile,
        //     header="=== Dump of Name object ===",
        //     footer="======= End of dump =======",
        //     )
        // throw biffh::XLRDError("Not a constant absolute reference to a single cell");
        throw std::runtime_error("Not a constant absolute reference to a single cell");
    };

    ////
    // This is a convenience method for the use case where the name
    // refers to one rectangular area in one worksheet.
    // @param clipped If true (the default), the returned rectangle is clipped
    // to fit in (0, sheet.nrows, 0, sheet.ncols) -- it is guaranteed that
    // 0 <= rowxlo <= rowxhi <= sheet.nrows and that the number of usable rows
    // in the area (which may be zero) is rowxhi - rowxlo; likewise for columns.
    // @return a tuple (sheet_object, rowxlo, rowxhi, colxlo, colxhi).
    // @throws XLRDError The name is not a constant absolute reference
    // to a single area in a single sheet.
    inline
    std::tuple<sheet::Sheet, int, int, int, int>
    area2d()
    {
        // auto res = this->result;
        // if (res) {
        //     // result should be an instance of the Operand class
        //     int kind = res->kind;
        //     // auto value = res->value;
        //     // if (kind == formula::oREF && value.size() == 1) {  // only 1 reference
        //     //     ref3d = value[0]
        //     //     if (0 <= ref3d.shtxlo == ref3d.shtxhi - 1) {  // only 1 usable sheet
        //     //         sh = self.book.sheet_by_index(ref3d.shtxlo)
        //     //         if (!clipped) {
        //     //             return std::make_tuple(
        //     //                 sh, ref3d.rowxlo, ref3d.rowxhi, ref3d.colxlo, ref3d.colxhi
        //     //             );
        //     //         }
        //     //         int rowxlo = min(ref3d.rowxlo, sh.nrows);
        //     //         int rowxhi = max(rowxlo, min(ref3d.rowxhi, sh.nrows));
        //     //         int colxlo = min(ref3d.colxlo, sh.ncols);
        //     //         int colxhi = max(colxlo, min(ref3d.colxhi, sh.ncols));
        //     //         assert 0 <= rowxlo <= rowxhi <= sh.nrows;
        //     //         assert 0 <= colxlo <= colxhi <= sh.ncols;
        //     //         return sh, rowxlo, rowxhi, colxlo, colxhi;
        //     //     }
        //     // }
        // }
        // throw XLRDError("Not a constant absolute reference to a single area in a single sheet");
        throw std::runtime_error("Not a constant absolute reference to a single area in a single sheet");
    };


};

class Book
: public formula::FormulaDelegate
, public sheet::SheetOwnerInterface
, public formatting::FormattingDelegate
{
public:
    ////
    // The number of worksheets present in the workbook file.
    // This information is available even when no sheets have yet been loaded.
    int nsheets;  // = 0

    ////
    // Which date system was in force when this file was last saved.<br />
    //    0 => 1900 system (the Excel for Windows default).<br />
    //    1 => 1904 system (the Excel for Macintosh default).<br />
    int datemode;  // = 0 // In case it's not specified in the file.

    ////
    // Version of BIFF (Binary Interchange File Format) used to create the file.
    // Latest is 8.0 (represented here as 80), introduced with Excel 97.
    // Earliest supported by this module: 2.0 (represented as 20).
    int biff_version = 0;
    ////
    // List containing a Name object for each NAME record in the workbook.
    // <br />  -- New in version 0.6.0
    std::vector<Name> name_obj_list;

    ////
    // An integer denoting the character set used for strings in this file.
    // For BIFF 8 and later, this will be 1200, meaning Unicode; more precisely, UTF_16_LE.
    // For earlier versions, this is used to derive the appropriate Python encoding
    // to be used to convert to Unicode.
    // Examples: 1252 -> 'cp1252', 10000 -> 'mac_roman'
    int codepage; // = None

    ////
    // The encoding that was derived from the codepage.
    std::string encoding; // = None

    ////
    // A tuple containing the (telephone system) country code for:<br />
    //    [0]: the user-interface setting when the file was created.<br />
    //    [1]: the regional settings.<br />
    // Example: (1, 61) meaning (USA, Australia).
    // This information may give a clue to the correct encoding for an unknown codepage.
    // For a long list of observed values, refer to the OpenOffice.org documentation for
    // the COUNTRY record.
    std::array<int, 2> countries;  // = (0, 0)

    ////
    // What (if anything) is recorded as the name of the last user to save the file.
    std::string user_name;  // = UNICODE_LITERAL('')

    ////
    // A list of Font class instances, each corresponding to a FONT record.
    // <br /> -- New in version 0.6.1
    std::vector<utils::any> font_list; // = []

    ////
    // A list of XF class instances, each corresponding to an XF record.
    // <br /> -- New in version 0.6.1
    std::vector<utils::any> xf_list;  // = []

    ////
    // A list of Format objects, each corresponding to a FORMAT record, in
    // the order that they appear in the input file.
    // It does <i>not</i> contain builtin formats.
    // If you are creating an output file using (for example) pyExcelerator,
    // use this list.
    // The collection to be used for all visual rendering purposes is format_map.
    // <br /> -- New in version 0.6.1
    std::vector<utils::any> format_list;  // = []

    ////
    // The mapping from XF.format_key to Format object.
    // <br /> -- New in version 0.6.1
    std::map<int, utils::any> format_map;  // = {}

    ////
    // This provides access via name to the extended format information for
    // both built-in styles and user-defined styles.<br />
    // It maps <i>name</i> to (<i>built_in</i>, <i>xf_index</i>), where:<br />
    // <i>name</i> is either the name of a user-defined style,
    // or the name of one of the built-in styles. Known built-in names are
    // Normal, RowLevel_1 to RowLevel_7,
    // ColLevel_1 to ColLevel_7, Comma, Currency, Percent, "Comma [0]",
    // "Currency [0]", Hyperlink, and "Followed Hyperlink".<br />
    // <i>built_in</i> 1 = built-in style, 0 = user-defined<br />
    // <i>xf_index</i> is an index into Book.xf_list.<br />
    // References: OOo docs s6.99 (STYLE record); Excel UI Format/Style
    // <br /> -- New in version 0.6.1; since 0.7.4, extracted only if
    // open_workbook(..., formatting_info=True)
    std::map<std::string, utils::any> style_name_map;  /// = {}

    ////
    // This provides definitions for colour indexes. Please refer to the
    // above section "The Palette; Colour Indexes" for an explanation
    // of how colours are represented in Excel.<br />
    // Colour indexes into the palette map into (red, green, blue) tuples.
    // "Magic" indexes e.g. 0x7FFF map to None.
    // <i>colour_map</i> is what you need if you want to render cells on screen or in a PDF
    // file. If you are writing an output XLS file, use <i>palette_record</i>.
    // <br /> -- New in version 0.6.1. Extracted only if open_workbook(..., formatting_info=True)
    std::map<int, utils::any> colour_map;  // = {}

    ////
    // If the user has changed any of the colours in the standard palette, the XLS
    // file will contain a PALETTE record with 56 (16 for Excel 4.0 and earlier)
    // RGB values in it, and this list will be e.g. [(r0, b0, g0), ..., (r55, b55, g55)].
    // Otherwise this list will be empty. This is what you need if you are
    // writing an output XLS file. If you want to render cells on screen or in a PDF
    // file, use colour_map.
    // <br /> -- New in version 0.6.1. Extracted only if open_workbook(..., formatting_info=True)
    std::vector<utils::any> palette_record;  // = []

    ////
    // Time in seconds to extract the XLS image as a contiguous string (or mmap equivalent).
    double load_time_stage_1;  // = -1.0

    ////
    // Time in seconds to parse the data from the contiguous string (or mmap equivalent).
    double load_time_stage_2;  // = -1.0

    ////
    // @return A list of all sheets in the book.
    // All sheets not already loaded will be loaded.
    std::vector<sheet::Sheet> sheets();

    ////
    // @param sheetx Sheet index in range(nsheets)
    // @return An object of the Sheet class
    sheet::Sheet sheet_by_index(int sheetx);

    ////
    // @param sheet_name Name of sheet required
    // @return An object of the Sheet class
    sheet::Sheet sheet_by_name(std::string sheet_name);

    ////
    // @return A list of the names of all the worksheets in the workbook file.
    // This information is available even when no sheets have yet been loaded.
    std::vector<std::string> sheet_names();

    ////
    // @param sheet_name_or_index Name or index of sheet enquired upon
    // @return true if sheet is loaded, false otherwise
    // <br />  -- New in version 0.7.1
    bool sheet_loaded(int sheet_index);
    bool sheet_loaded(std::string sheet_name);

    ////
    // @param sheet_name_or_index Name or index of sheet to be unloaded.
    // <br />  -- New in version 0.7.1
    void unload_sheet(int sheet_index);
    void unload_sheet(std::string sheet_name);
        
    ////
    // This method has a dual purpose. You can call it to release
    // memory-consuming objects and (possibly) a memory-mapped file
    // (mmap.mmap object) when you have finished loading sheets in
    // on_demand mode, but still require the Book object to examine the
    // loaded sheets. It is also called automatically (a) when open_workbook
    // raises an exception and (b) if you are using a "with" statement, when 
    // the "with" block is exited. Calling this method multiple times on the 
    // same object has no ill effect.
    void release_resources();
    
    ////
    // A mapping from (lower_case_name, scope) to a single Name object.
    // <br />  -- New in version 0.6.0
    std::map<std::string, utils::any> name_and_scope_map;

    ////
    // A mapping from lower_case_name to a list of Name objects. The list is
    // sorted in scope order. Typically there will be one item (of global scope)
    // in the list.
    // <br />  -- New in version 0.6.0
    std::map<std::string, Name> name_map; // = {}

    int logfile = 0;
    int verbosity = 0;
    int use_mmap = 0;
    std::string encoding_override;
    int formatting_info = 0;
    int on_demand = 0;
    int ragged_rows = 0;
    std::map<int, int> _xf_index_to_xl_type_map;
    int base;
    int _position;
    std::vector<uint8_t> filestr;
    std::vector<uint8_t> mem;
    size_t stream_len;

    std::vector<sheet::Sheet> _sheet_list;
    std::vector<std::string> _sheet_names;
    std::vector<int> _sheet_visibility;
    std::vector<int> _sh_abs_posn;

    std::string raw_user_name;
    int builtinfmtcount;
    int _supbook_count;
    std::vector<int> _externsheet_type_b57;
    std::vector<std::string> _extnsht_name_from_num;
    std::map<std::string, int> _sheet_num_from_name;
    int _extnsht_count;
    std::vector<int> _supbook_types;
    int _resources_released;
    std::vector<std::string> addin_func_names;

    inline
    Book() {
        this->_sheet_list = {};
        this->_sheet_names = {};
        this->_sheet_visibility = {};  // from BOUNDSHEET record
        this->nsheets = 0;
        this->_sh_abs_posn = {};  // sheet's absolute position in the stream
        //this->_sharedstrings = {};
        //this->_rich_text_runlist_map = {};
        this->raw_user_name = "";
        //this->_sheethdr_count = 0;  // BIFF 4W only
        this->builtinfmtcount = -1;  // unknown as yet. BIFF 3, 4S, 4W
        this->initialise_format_info();
        //this->_all_sheets_count = 0;  // includes macro & VBA sheets
        this->_supbook_count = 0;
        this->_supbook_locals_inx = -1;
        this->_supbook_addins_inx = -1;
        this->_all_sheets_map = {};  // maps an all_sheets index to a calc-sheets index (or -1)
        this->_externsheet_info = {};
        this->_externsheet_type_b57 = {};
        this->_extnsht_name_from_num = {};
        this->_sheet_num_from_name = {};
        this->_extnsht_count = 0;
        this->_supbook_types = {};
        this->_resources_released = 0;
        this->addin_func_names = {};
        this->name_obj_list = {};
        this->colour_map = {};
        this->palette_record = {};
        this->xf_list = {};
        this->style_name_map = {};
        this->mem = {};
        this->filestr = {};
    }

    inline
    void biff2_8_load(std::vector<uint8_t> file_contents)
    {
        // DEBUG = 0
        this->logfile = 0;
        this->verbosity = 0;
        this->use_mmap = 0;
        this->encoding_override = "";
        this->formatting_info = 0;
        this->on_demand = 0;
        this->ragged_rows = 0;

        this->filestr = file_contents;
        this->stream_len = file_contents.size();

        this->base = 0;
        this->mem.clear();
        if (!utils::equals(utils::slice(this->filestr, 0, 8), compdoc::SIGNATURE)) {
            // got this one at the antique store
            this->mem = this->filestr;
        }
        else {
            auto cd = compdoc::CompDoc(this->filestr);
            if (USE_FANCY_CD) {
                std::tie(this->mem, this->base, this->stream_len) = cd.locate_named_stream("Workbook");
                if (!this->mem.empty()) {
                    std::tie(this->mem, this->base, this->stream_len) = cd.locate_named_stream("Book");
                }
                if (this->mem.empty()) {
                    throw XLRDError("Can't find workbook in OLE2 compound document");
                }
            }
            else {
                this->mem = cd.get_named_stream("Workbook");
                if (this->mem.empty()) {
                    this->mem = cd.get_named_stream("Book");
                }
                if (this->mem.empty()) {
                    throw XLRDError("Can't find workbook in OLE2 compound document");
                }
                this->stream_len = this->mem.size();
            }
        }
        this->_position = this->base;
        if (DEBUG) {
            pprint("mem: %s, base: %d, len: %d", this->mem, this->base, this->stream_len);
        }
    }

    inline
    void initialise_format_info();

    inline
    int get2bytes();

    inline
    std::tuple<int, int, std::vector<uint8_t>>
    get_record_parts() {
        int pos = this->_position;
        auto& mem = this->mem;
        int code = utils::as_uint16(mem, pos);
        int length = utils::as_uint16(mem, pos+2);
        pos += 4;
        auto data = utils::slice(mem, pos, pos+length);
        this->_position = pos + length;
        return std::make_tuple(code, length, data);
    }

    inline
    std::tuple<int, int, std::vector<uint8_t>>
    get_record_parts_conditional(int reqd_record);

    inline
    sheet::Sheet*
    get_sheet(int sh_number, bool update_pos=true);

    void get_sheets();

    void fake_globals_get_sheet(); // for BIFF 4.0 and earlier

    void handle_boundsheet(std::vector<uint8_t> data);

    void handle_builtinfmtcount(std::vector<uint8_t>& data);

    virtual
    std::string derive_encoding() {
        if (!this->encoding_override.empty()) {
            this->encoding = this->encoding_override;
        } else if (this->codepage == 0) {
            if (this->biff_version < 80) {
                pprint(
                    "*** No CODEPAGE record, no encoding_override: will use 'ascii'\n");
                this->encoding = "ascii";
            } else {
                this->codepage = 1200; // utf16le
                if (this->verbosity >= 2) {
                    pprint(
                      "*** No CODEPAGE record; assuming 1200 (utf_16_le)\n");
                }
            }
        } else {
            int codepage = this->codepage;
            std::string encoding;
            if (utils::haskey(encoding_from_codepage, codepage)) {
                encoding = encoding_from_codepage.at(codepage);
            } else if (300 <= codepage && codepage <= 1999) {
                encoding = format("cp%d", codepage);
            } else {
                encoding = format("unknown_codepage_%d", codepage);
            }
            if (DEBUG or (this->verbosity and encoding != this->encoding)) {
                pprint("CODEPAGE: codepage %d -> encoding %s\n", codepage, encoding);
            }
            this->encoding = encoding;
        }
        if (this->codepage != 1200) { // utf_16_le
            // If we don't have a codec that can decode ASCII into Unicode,
            // we're well & truly stuffed -- let the punter know ASAP.
            //try:
            //    _unused = unicode(b'trial', this->encoding)
            //except BaseException as e:
            //    fprintf(this->logfile,
            //        "ERROR *** codepage %r -> encoding %r -> %s: %s\n",
            //        this->codepage, this->encoding, type(e).__name__.split(".")[-1], e)
            //    raise
        }
        if (!this->raw_user_name.empty()) {
            auto strg = unpack_string(this->user_name, 0, this->encoding, 1);
            strg = strutil::rtrim(strg);
            // if DEBUG:
            //     print "CODEPAGE: user name decoded from %r to %r" % (this->user_name, strg)
            this->user_name = strg;
            this->raw_user_name = "";
        }
        return this->encoding;
    }

    inline
    void handle_codepage(const std::vector<uint8_t>& data) {
        int codepage = utils::as_uint16(data, 0);
        this->codepage = codepage;
        this->derive_encoding();
    }

    inline
    void handle_country(const std::vector<uint8_t>&  data) {
        int country0 = utils::as_uint16(data, 0);
        int country1 = utils::as_uint16(data, 2);
        if (self.verbosity) {
            // pprint("Countries:%s", countries);
        }
        // Note: in BIFF7 and earlier, country record was put (redundantly?) in each worksheet.
        // ASSERT(self.countries == (0, 0) or self.countries == countries);
        this->countries[0] = country0;
        this->countries[1] = country1;
    }

    inline void
    handle_datemode(const vector<u8>& data) {
        int datemode = as_uint16(data, 0);
        if (DEBUG or this->verbosity) {
            pprint("DATEMODE: datemode %r\n", datemode);
        }
        assert(datemode==0 or datemode==1);
        this->datemode = datemode;
    }

    inline void
    handle_externname(const vector<u8>& data) {
        int blah = DEBUG or self.verbosity >= 2;
        if (this->biff_version >= 80) {
            int option_flags = as_uint16(data, 0);
            int other_info = as_int32(data, 2);
            int pos = 6;
            std:string name;
            std::tie(name, pos) = unpack_unicode_update_pos(
                                      data, pos, 1);
            auto extra = slice(data, pos, 0);
            if (this->_supbook_types.back() == SUPBOOK_ADDIN) {
                this->addin_func_names.push_back(name);
            }
            if (blah) {
                pprint(
                    "EXTERNNAME: sbktype=%d oflags=0x%04x oinfo=0x%08x name=%r extra=%r\n",
                    this->_supbook_types.back(), option_flags, other_info, name, extra);
            }
        }
    }
    
    inline void
    handle_externsheet(const vector<u8>& data) {
        this->derive_encoding() // in case CODEPAGE record missing/out of order/wrong
        this->_extnsht_count += 1; // for use as a 1-based index
        int blah1 = DEBUG or self.verbosity >= 1;
        int blah2 = DEBUG or self.verbosity >= 2;
        if (this->biff_version >= 80) {
            int num_refs = as_uint16(data, 0);
            int bytes_reqd = num_refs * 6 + 2;
            while len(data) < bytes_reqd:
                if blah1:
                    fprintf(
                        self.logfile,
                        "INFO: EXTERNSHEET needs %d bytes, have %d\n",
                        bytes_reqd, len(data),
                        )
                code2, length2, data2 = self.get_record_parts()
                if code2 != XL_CONTINUE:
                    raise XLRDError("Missing CONTINUE after EXTERNSHEET record")
                data += data2
            pos = 2
            for k in xrange(num_refs):
                info = unpack("<HHH", data[pos:pos+6])
                ref_recordx, ref_first_sheetx, ref_last_sheetx = info
                self._externsheet_info.append(info)
                pos += 6
                if blah2:
                    fprintf(
                        self.logfile,
                        "EXTERNSHEET(b8): k = %2d, record = %2d, first_sheet = %5d, last sheet = %5d\n",
                        k, ref_recordx, ref_first_sheetx, ref_last_sheetx,
                        )
        else:
            nc, ty = unpack("<BB", data[:2])
            if blah2:
                print("EXTERNSHEET(b7-):", file=self.logfile)
                hex_char_dump(data, 0, len(data), fout=self.logfile)
                msg = {
                    1: "Encoded URL",
                    2: "Current sheet!!",
                    3: "Specific sheet in own doc't",
                    4: "Nonspecific sheet in own doc't!!",
                    }.get(ty, "Not encoded")
                print("   %3d chars, type is %d (%s)" % (nc, ty, msg), file=self.logfile)
            if ty == 3:
                sheet_name = unicode(data[2:nc+2], self.encoding)
                self._extnsht_name_from_num[self._extnsht_count] = sheet_name
                if blah2: print(self._extnsht_name_from_num, file=self.logfile)
            if not (1 <= ty <= 4):
                ty = 0
            self._externsheet_type_b57.append(ty)

    def handle_filepass(self, data):
        if self.verbosity >= 2:
            logf = self.logfile
            fprintf(logf, "FILEPASS:\n")
            hex_char_dump(data, 0, len(data), base=0, fout=logf)
            if self.biff_version >= 80:
                kind1, = unpack('<H', data[:2])
                if kind1 == 0: // weak XOR encryption
                    key, hash_value = unpack('<HH', data[2:])
                    fprintf(logf,
                        'weak XOR: key=0x%04x hash=0x%04x\n',
                        key, hash_value)
                elif kind1 == 1:
                    kind2, = unpack('<H', data[4:6])
                    if kind2 == 1: // BIFF8 standard encryption
                        caption = "BIFF8 std"
                    elif kind2 == 2:
                        caption = "BIFF8 strong"
                    else:
                        caption = "** UNKNOWN ENCRYPTION METHOD **"
                    fprintf(logf, "%s\n", caption)
        raise XLRDError("Workbook is encrypted")

    def handle_name(self, data):
        blah = DEBUG or self.verbosity >= 2
        bv = self.biff_version
        if bv < 50:
            return
        self.derive_encoding()
        // print
        // hex_char_dump(data, 0, len(data), fout=self.logfile)
        (
        option_flags, kb_shortcut, name_len, fmla_len, extsht_index, sheet_index,
        menu_text_len, description_text_len, help_topic_text_len, status_bar_text_len,
        ) = unpack("<HBBHHH4B", data[0:14])
        nobj = Name()
        nobj.book = self ////// CIRCULAR //////
        name_index = len(self.name_obj_list)
        nobj.name_index = name_index
        self.name_obj_list.append(nobj)
        nobj.option_flags = option_flags
        for attr, mask, nshift in (
            ('hidden', 1, 0),
            ('func', 2, 1),
            ('vbasic', 4, 2),
            ('macro', 8, 3),
            ('complex', 0x10, 4),
            ('builtin', 0x20, 5),
            ('funcgroup', 0xFC0, 6),
            ('binary', 0x1000, 12),
            ):
            setattr(nobj, attr, (option_flags & mask) >> nshift)

        macro_flag = " M"[nobj.macro]
        if bv < 80:
            internal_name, pos = unpack_string_update_pos(data, 14, self.encoding, known_len=name_len)
        else:
            internal_name, pos = unpack_unicode_update_pos(data, 14, known_len=name_len)
        nobj.extn_sheet_num = extsht_index
        nobj.excel_sheet_index = sheet_index
        nobj.scope = None // patched up in the names_epilogue() method
        if blah:
            fprintf(
                self.logfile,
                "NAME[%d]:%s oflags=%d, name_len=%d, fmla_len=%d, extsht_index=%d, sheet_index=%d, name=%r\n",
                name_index, macro_flag, option_flags, name_len,
                fmla_len, extsht_index, sheet_index, internal_name)
        name = internal_name
        if nobj.builtin:
            name = builtin_name_from_code.get(name, "??Unknown??")
            if blah: print("    builtin: %s" % name, file=self.logfile)
        nobj.name = name
        nobj.raw_formula = data[pos:]
        nobj.basic_formula_len = fmla_len;
        nobj.evaluated = 0
        if blah:
            nobj.dump(
                self.logfile,
                header="--- handle_name: name[%d] ---" % name_index,
                footer="-------------------",
                )

    def names_epilogue(self):
        blah = self.verbosity >= 2
        f = self.logfile
        if blah:
            print("+++++ names_epilogue +++++", file=f)
            print("_all_sheets_map", REPR(self._all_sheets_map), file=f)
            print("_extnsht_name_from_num", REPR(self._extnsht_name_from_num), file=f)
            print("_sheet_num_from_name", REPR(self._sheet_num_from_name), file=f)
        num_names = len(self.name_obj_list)
        for namex in range(num_names):
            nobj = self.name_obj_list[namex]
            // Convert from excel_sheet_index to scope.
            // This is done here because in BIFF7 and earlier, the
            // BOUNDSHEET records (from which _all_sheets_map is derived)
            // come after the NAME records.
            if self.biff_version >= 80:
                sheet_index = nobj.excel_sheet_index
                if sheet_index == 0:
                    intl_sheet_index = -1 // global
                elif 1 <= sheet_index <= len(self._all_sheets_map):
                    intl_sheet_index = self._all_sheets_map[sheet_index-1]
                    if intl_sheet_index == -1: // maps to a macro or VBA sheet
                        intl_sheet_index = -2 // valid sheet reference but not useful
                else:
                    // huh?
                    intl_sheet_index = -3 // invalid
            elif 50 <= self.biff_version <= 70:
                sheet_index = nobj.extn_sheet_num
                if sheet_index == 0:
                    intl_sheet_index = -1 // global
                else:
                    sheet_name = self._extnsht_name_from_num[sheet_index]
                    intl_sheet_index = self._sheet_num_from_name.get(sheet_name, -2)
            nobj.scope = intl_sheet_index

        for namex in range(num_names):
            nobj = self.name_obj_list[namex]
            // Parse the formula ...
            if nobj.macro or nobj.binary: continue
            if nobj.evaluated: continue
            evaluate_name_formula(self, nobj, namex, blah=blah)

        if self.verbosity >= 2:
            print("---------- name object dump ----------", file=f)
            for namex in range(num_names):
                nobj = self.name_obj_list[namex]
                nobj.dump(f, header="--- name[%d] ---" % namex)
            print("--------------------------------------", file=f)
        //
        // Build some dicts for access to the name objects
        //
        name_and_scope_map = {} // (name.lower(), scope): Name_object
        name_map = {}           // name.lower() : list of Name_objects (sorted in scope order)
        for namex in range(num_names):
            nobj = self.name_obj_list[namex]
            name_lcase = nobj.name.lower()
            key = (name_lcase, nobj.scope)
            if key in name_and_scope_map and self.verbosity:
                fprintf(f, 'Duplicate entry %r in name_and_scope_map\n', key)
            name_and_scope_map[key] = nobj
            sort_data = (nobj.scope, namex, nobj)
            // namex (a temp unique ID) ensures the Name objects will not
            // be compared (fatal in py3)
            if name_lcase in name_map:
                name_map[name_lcase].append(sort_data)
            else:
                name_map[name_lcase] = [sort_data]
        for key in name_map.keys():
            alist = name_map[key]
            alist.sort()
            name_map[key] = [x[2] for x in alist]
        self.name_and_scope_map = name_and_scope_map
        self.name_map = name_map

    def handle_obj(self, data):
        // Not doing much handling at all.
        // Worrying about embedded (BOF ... EOF) substreams is done elsewhere.
        // DEBUG = 1
        obj_type, obj_id = unpack('<HI', data[4:10])
        // if DEBUG: print "---> handle_obj type=%d id=0x%08x" % (obj_type, obj_id)

    def handle_supbook(self, data):
        // aka EXTERNALBOOK in OOo docs
        self._supbook_types.append(None)
        blah = DEBUG or self.verbosity >= 2
        if blah:
            print("SUPBOOK:", file=self.logfile)
            hex_char_dump(data, 0, len(data), fout=self.logfile)
        num_sheets = unpack("<H", data[0:2])[0]
        if blah: print("num_sheets = %d" % num_sheets, file=self.logfile)
        sbn = self._supbook_count
        self._supbook_count += 1
        if data[2:4] == b"\x01\x04":
            self._supbook_types[-1] = SUPBOOK_INTERNAL
            self._supbook_locals_inx = self._supbook_count - 1
            if blah:
                print("SUPBOOK[%d]: internal 3D refs; %d sheets" % (sbn, num_sheets), file=self.logfile)
                print("    _all_sheets_map", self._all_sheets_map, file=self.logfile)
            return
        if data[0:4] == b"\x01\x00\x01\x3A":
            self._supbook_types[-1] = SUPBOOK_ADDIN
            self._supbook_addins_inx = self._supbook_count - 1
            if blah: print("SUPBOOK[%d]: add-in functions" % sbn, file=self.logfile)
            return
        url, pos = unpack_unicode_update_pos(data, 2, lenlen=2)
        if num_sheets == 0:
            self._supbook_types[-1] = SUPBOOK_DDEOLE
            if blah: fprintf(self.logfile, "SUPBOOK[%d]: DDE/OLE document = %r\n", sbn, url)
            return
        self._supbook_types[-1] = SUPBOOK_EXTERNAL
        if blah: fprintf(self.logfile, "SUPBOOK[%d]: url = %r\n", sbn, url)
        sheet_names = []
        for x in range(num_sheets):
            try:
                shname, pos = unpack_unicode_update_pos(data, pos, lenlen=2)
            except struct.error:
                // //////// FIX ME ////////
                // Should implement handling of CONTINUE record(s) ...
                if self.verbosity:
                    print((
                        "*** WARNING: unpack failure in sheet %d of %d in SUPBOOK record for file %r" 
                        % (x, num_sheets, url)
                        ), file=self.logfile)
                break
            sheet_names.append(shname)
            if blah: fprintf(self.logfile, "  sheetx=%d namelen=%d name=%r (next pos=%d)\n", x, len(shname), shname, pos)

    def handle_sheethdr(self, data):
        // This a BIFF 4W special.
        // The SHEETHDR record is followed by a (BOF ... EOF) substream containing
        // a worksheet.
        // DEBUG = 1
        self.derive_encoding()
        sheet_len = unpack('<i', data[:4])[0]
        sheet_name = unpack_string(data, 4, self.encoding, lenlen=1)
        sheetno = self._sheethdr_count
        assert sheet_name == self._sheet_names[sheetno]
        self._sheethdr_count += 1
        BOF_posn = self._position
        posn = BOF_posn - 4 - len(data)
        if DEBUG: fprintf(self.logfile, 'SHEETHDR %d at posn %d: len=%d name=%r\n', sheetno, posn, sheet_len, sheet_name)
        self.initialise_format_info()
        if DEBUG: print('SHEETHDR: xf epilogue flag is %d' % self._xf_epilogue_done, file=self.logfile)
        self._sheet_list.append(None) // get_sheet updates _sheet_list but needs a None beforehand
        self.get_sheet(sheetno, update_pos=False)
        if DEBUG: print('SHEETHDR: posn after get_sheet() =', self._position, file=self.logfile)
        self._position = BOF_posn + sheet_len

    def handle_sheetsoffset(self, data):
        // DEBUG = 0
        posn = unpack('<i', data)[0]
        if DEBUG: print('SHEETSOFFSET:', posn, file=self.logfile)
        self._sheetsoffset = posn

    def handle_sst(self, data):
        // DEBUG = 1
        if DEBUG:
            print("SST Processing", file=self.logfile)
            t0 = time.time()
        nbt = len(data)
        strlist = [data]
        uniquestrings = unpack('<i', data[4:8])[0]
        if DEBUG  or self.verbosity >= 2:
            fprintf(self.logfile, "SST: unique strings: %d\n", uniquestrings)
        while 1:
            code, nb, data = self.get_record_parts_conditional(XL_CONTINUE)
            if code is None:
                break
            nbt += nb
            if DEBUG >= 2:
                fprintf(self.logfile, "CONTINUE: adding %d bytes to SST -> %d\n", nb, nbt)
            strlist.append(data)
        self._sharedstrings, rt_runlist = unpack_SST_table(strlist, uniquestrings)
        if self.formatting_info:
            self._rich_text_runlist_map = rt_runlist        
        if DEBUG:
            t1 = time.time()
            print("SST processing took %.2f seconds" % (t1 - t0, ), file=self.logfile)

    def handle_writeaccess(self, data):
        DEBUG = 0
        if self.biff_version < 80:
            if not self.encoding:
                self.raw_user_name = True
                self.user_name = data
                return
            strg = unpack_string(data, 0, self.encoding, lenlen=1)
        else:
            strg = unpack_unicode(data, 0, lenlen=2)
        if DEBUG: fprintf(self.logfile, "WRITEACCESS: %d bytes; raw=%s %r\n", len(data), self.raw_user_name, strg)
        strg = strg.rstrip()
        self.user_name = strg
*/

    inline
    void xf_epilogue() {
        formatting::xf_epilogue(this);
    }

    inline
    void parse_globals()
    {
        // DEBUG = 0
        // no need to position, just start reading (after the BOF)
        formatting::initialise_book(this);
        while (1) {
            int rc, length;
            std::vector<uint8_t> data;
            std::tie(rc, length, data) = this->get_record_parts();
            if (DEBUG){
                pprint("parse_globals: record code is 0x%04x", rc);
            }
            if (rc == biffh::XL_SST) {
                this->handle_sst(data);
            } else if (rc == biffh::XL_FONT || rc == biffh::XL_FONT_B3B4) {
                this->handle_font(data);
            } else if (rc == biffh::XL_FORMAT) { // biffh::XL_FORMAT2 is BIFF <= 3.0, can't appear in globals
                this->handle_format(data);
            } else if (rc == biffh::XL_XF) {
                this->handle_xf(data);
            } else if (rc ==  biffh::XL_BOUNDSHEET) {
                this->handle_boundsheet(data);
            } else if (rc == biffh::XL_DATEMODE) {
                this->handle_datemode(data);
            } else if (rc == biffh::XL_CODEPAGE) {
                this->handle_codepage(data);
            } else if (rc == biffh::XL_COUNTRY) {
                this->handle_country(data);
            } else if (rc == biffh::XL_EXTERNNAME) {
                this->handle_externname(data);
            } else if (rc == biffh::XL_EXTERNSHEET) {
                this->handle_externsheet(data);
            } else if (rc == biffh::XL_FILEPASS) {
                this->handle_filepass(data);
            } else if (rc == biffh::XL_WRITEACCESS) {
                this->handle_writeaccess(data);
            } else if (rc == biffh::XL_SHEETSOFFSET) {
                this->handle_sheetsoffset(data);
            } else if (rc == biffh::XL_SHEETHDR) {
                this->handle_sheethdr(data);
            } else if (rc == biffh::XL_SUPBOOK) {
                this->handle_supbook(data);
            } else if (rc == biffh::XL_NAME) {
                this->handle_name(data);
            } else if (rc == biffh::XL_PALETTE) {
                this->handle_palette(data);
            } else if (rc == biffh::XL_STYLE) {
                this->handle_style(data);
            } else if ((rc & 0xff) == 9 && this->verbosity) {
                pprint(
                    "*** Unexpected BOF at posn %d: 0x%04x len=%d data=%s\n",
                    this->_position - length - 4, rc, length, utils::str::repr(data));
            } else if (rc == biffh::XL_EOF) {
                this->xf_epilogue();
                this->names_epilogue();
                this->palette_epilogue();
                if (this->encoding.empty()) {
                    this->derive_encoding();
                }
                if (this->biff_version == 45) {
                    // DEBUG = 0
                    if (DEBUG) pprint("global EOF: position=%d", self._position);
                    // if DEBUG:
                    //     pos = self._position - 4
                    //     print repr(self.mem[pos:pos+40])
                }
                return;
            }
            else {
                // if DEBUG:
                //     print >> self.logfile, "parse_globals: ignoring record code 0x%04x" % rc
            }
        }
    }

    inline
    std::vector<uint8_t>
    read(int pos, int length) {
        auto data = utils::slice(this->mem, pos, pos+length);
        self._position = pos + data.size();
        return data;
    }

    inline
    void getbof(int rqd_stream) {
/*
        // DEBUG = 1
        // if DEBUG: print >> self.logfile, "getbof(): position", self._position
        if DEBUG: print("reqd: 0x%04x" % rqd_stream, file=self.logfile)
        def bof_error(msg):
            raise XLRDError('Unsupported format, or corrupt file: ' + msg)
        savpos = self._position
        opcode = self.get2bytes()
        if opcode == MY_EOF:
            bof_error('Expected BOF record; met end of file')
        if opcode not in bofcodes:
            bof_error('Expected BOF record; found %r' % self.mem[savpos:savpos+8])
        length = self.get2bytes()
        if length == MY_EOF:
            bof_error('Incomplete BOF record[1]; met end of file')
        if not (4 <= length <= 20):
            bof_error(
                'Invalid length (%d) for BOF record type 0x%04x'
                % (length, opcode))
        padding = b'\0' * max(0, boflen[opcode] - length)
        data = self.read(self._position, length);
        if DEBUG: fprintf(self.logfile, "\ngetbof(): data=%r\n", data)
        if len(data) < length:
            bof_error('Incomplete BOF record[2]; met end of file')
        data += padding
        version1 = opcode >> 8
        version2, streamtype = unpack('<HH', data[0:4])
        if DEBUG:
            print("getbof(): op=0x%04x version2=0x%04x streamtype=0x%04x" \
                % (opcode, version2, streamtype), file=self.logfile)
        bof_offset = self._position - 4 - length
        if DEBUG:
            print("getbof(): BOF found at offset %d; savpos=%d" \
                % (bof_offset, savpos), file=self.logfile)
        version = build = year = 0
        if version1 == 0x08:
            build, year = unpack('<HH', data[4:8])
            if version2 == 0x0600:
                version = 80
            elif version2 == 0x0500:
                if year < 1994 or build in (2412, 3218, 3321):
                    version = 50
                else:
                    version = 70
            else:
                // dodgy one, created by a 3rd-party tool
                version = {
                    0x0000: 21,
                    0x0007: 21,
                    0x0200: 21,
                    0x0300: 30,
                    0x0400: 40,
                    }.get(version2, 0)
        elif version1 in (0x04, 0x02, 0x00):
            version = {0x04: 40, 0x02: 30, 0x00: 21}[version1]

        if version == 40 and streamtype == XL_WORKBOOK_GLOBALS_4W:
            version = 45 // i.e. 4W

        if DEBUG or self.verbosity >= 2:
            print("BOF: op=0x%04x vers=0x%04x stream=0x%04x buildid=%d buildyr=%d -> BIFF%d" \
                % (opcode, version2, streamtype, build, year, version), file=self.logfile)
        got_globals = streamtype == XL_WORKBOOK_GLOBALS or (
            version == 45 and streamtype == XL_WORKBOOK_GLOBALS_4W)
        if (rqd_stream == XL_WORKBOOK_GLOBALS and got_globals) or streamtype == rqd_stream:
            return version
        if version < 50 and streamtype == XL_WORKSHEET:
            return version
        if version >= 50 and streamtype == 0x0100:
            bof_error("Workspace file -- no spreadsheet data")
        bof_error(
            'BOF not workbook/worksheet: op=0x%04x vers=0x%04x strm=0x%04x build=%d year=%d -> BIFF%d' \
            % (opcode, version2, streamtype, build, year, version)
            )
*/
    }
};


inline
Book open_workbook_xls(std::vector<uint8_t> file_contents)
{
    // if TOGGLE_GC:
    //     orig_gc_enabled = gc.isenabled()
    //     if orig_gc_enabled:
    //         gc.disable()
    Book bk = Book();
    try {
        bk.biff2_8_load(file_contents);
        int biff_version = bk.getbof(biffh::XL_WORKBOOK_GLOBALS);
        if (biff_version == 0) {
            throw XLRDError("Can't determine file's BIFF version");
        }
        if (utils::indexof(SUPPORTED_VERSIONS, biff_version) == -1) {
            throw XLRDError(utils::str::format(
                "BIFF version %s is not supported"
                , biffh::biff_text_from_num[biff_version]
            ));
        }
        bk.biff_version = biff_version;
        if (biff_version <= 40) {
            // no workbook globals, only 1 worksheet
            pprint(
                "*** WARNING: on_demand is not supported for this Excel version.\n"
                "*** Setting on_demand to False.\n");
            bk.on_demand = false;
            bk.fake_globals_get_sheet();
        }
        else if (biff_version == 45) {
            // worksheet(s) embedded in global stream
            bk.parse_globals();
            pprint("*** WARNING: on_demand is not supported for this Excel version.\n"
                          "*** Setting on_demand to False.\n");
            bk.on_demand = false;
        }
        else {
            bk.parse_globals();
            int len = bk._sheet_names.size();
            bk._sheet_list.clear();
            bk.get_sheets();
        }
        bk.nsheets = bk._sheet_list.size();
        if (biff_version == 45 && bk.nsheets > 1) {
            pprint(
                "*** WARNING: Excel 4.0 workbook (.XLW) file contains %d worksheets.\n"
                "*** Book-level data will be that of the last worksheet.\n",
                bk.nsheets
            );
        }
        // bk.load_time_stage_2 = t2 - t1;
    } catch(std::exception exc) {
        bk.release_resources();
        throw;
    }
    bk.release_resources();
    return bk;
}


}
}
