#pragma once
// -*- coding: cp1252 -*-

////
// <p> Portions copyright � 2005-2013 Stephen John Machin, Lingfo Pty Ltd</p>
// <p>This module is part of the xlrd package, which is released under a BSD-style licence.</p>
////

// 2010-04-25 SJM fix zoom factors cooking logic
// 2010-04-15 CW  r4253 fix zoom factors cooking logic
// 2010-04-09 CW  r4248 add a flag so xlutils knows whether or not to write a PANE record
// 2010-03-29 SJM Fixed bug in adding new empty rows in put_cell_ragged
// 2010-03-28 SJM Tailored put_cell method for each of ragged_rows=False (fixed speed regression) and =True (faster)
// 2010-03-25 CW  r4236 Slight refactoring to remove method calls
// 2010-03-25 CW  r4235 Collapse expand_cells into put_cell and enhance the raggedness. This should save even more memory!
// 2010-03-25 CW  r4234 remove duplicate chunks for extend_cells; refactor to remove put_number_cell and put_blank_cell which essentially duplicated the code of put_cell
// 2010-03-10 SJM r4222 Added reading of the PANE record.
// 2010-03-10 SJM r4221 Preliminary work on "cooked" mag factors; use at own peril
// 2010-03-01 SJM Reading SCL record
// 2010-03-01 SJM Added ragged_rows functionality
// 2009-08-23 SJM Reduced CPU time taken by parsing MULBLANK records.
// 2009-08-18 SJM Used __slots__ and sharing to reduce memory consumed by Rowinfo instances
// 2009-05-31 SJM Fixed problem with no CODEPAGE record on extremely minimal BIFF2.x 3rd-party file
// 2009-04-27 SJM Integrated on_demand patch by Armando Serrano Lombillo
// 2008-02-09 SJM Excel 2.0: build XFs on the fly from cell attributes
// 2007-12-04 SJM Added support for Excel 2.x (BIFF2) files.
// 2007-10-11 SJM Added missing entry for blank cell type to ctype_text
// 2007-07-11 SJM Allow for BIFF2/3-style FORMAT record in BIFF4/8 file
// 2007-04-22 SJM Remove experimental "trimming" facility.

#include <vector>

#include "./biffh.h"  // __all__
#include "./formula.h"  // dump_formula, decompile_formula, rangename2d, FMLA_TYPE_CELL, FMLA_TYPE_SHARED
#include "./formatting.h"  // nearest_colour_index, Format

namespace xlrd {
namespace sheet {

using u8vec = std::vector<uint8_t>;

const auto& FUN = biffh::FUN;
const auto& FDT = biffh::FDT;
const auto& FNU = biffh::FNU;
const auto& FGE = biffh::FGE;
const auto& FTX = biffh::FTX;
const auto& XL_CELL_EMPTY = biffh::XL_CELL_EMPTY;
const auto& XL_CELL_TEXT = biffh::XL_CELL_TEXT;
const auto& XL_CELL_NUMBER = biffh::XL_CELL_NUMBER;
const auto& XL_CELL_DATE = biffh::XL_CELL_DATE;
const auto& XL_CELL_BOOLEAN = biffh::XL_CELL_BOOLEAN;
const auto& XL_CELL_ERROR = biffh::XL_CELL_ERROR;
const auto& XL_CELL_BLANK = biffh::XL_CELL_BLANK;


const int DEBUG = 0;
const int OBJ_MSO_DEBUG = 0;

const std::vector<std::pair<std::string, int>>
_WINDOW2_options = {
    // Attribute names and initial values to use in case
    // a WINDOW2 record is not written.
    {"show_formulas", 0},
    {"show_grid_lines", 1},
    {"show_sheet_headers", 1},
    {"panes_are_frozen", 0},
    {"show_zero_values", 1},
    {"automatic_grid_line_colour", 1},
    {"columns_from_right_to_left", 0},
    {"show_outline_symbols", 1},
    {"remove_splits_if_pane_freeze_is_removed", 0},
    // Multiple sheets can be selected, but only one can be active
    // {hold down Ctrl and click multiple tabs in the file in OOo}
    {"sheet_selected", 0},
    // "sheet_visible" should really be called "sheet_active"
    // and is 1 when this sheet is the sheet displayed when the file
    // is open. More than likely only one sheet should ever be set as
    // visible.
    // This would correspond to the Book's sheet_active attribute, but
    // that doesn't exist as WINDOW1 records aren't currently processed.
    // The real thing is the visibility attribute from the BOUNDSHEET record.
    {"sheet_visible", 0},
    {"show_in_page_break_preview", 0},
};

class Cell;

class SheetOwnerInterface {
public:
    int biff_version = 0;
    int verbosity;
    int formatting_info;
    int ragged_rows;
    std::map<int, int>* _xf_index_to_xl_type_map;
    int _maxdatarowx = -1; // highest rowx containing a non-empty cell
    int _maxdatacolx = -1; // highest colx containing a non-empty cell
    int _dimnrows = 0; // as per DIMENSIONS record
    int _dimncols = 0;
    std::vector<utils::any> _cell_values;
    std::vector<utils::any> _cell_types;
    std::vector<int> _cell_xf_indexes;
    std::vector<int> _xf_index_stats;
    std::vector<int> _sheet_visibility;
};


////
// <p>Contains the data for one worksheet.</p>
//
// <p>In the cell access functions, "rowx" is a row index, counting from zero, and "colx" is a
// column index, counting from zero.
// Negative values for row/column indexes and slice positions are supported in the expected fashion.</p>
//
// <p>For information about cell types and cell values, refer to the documentation of the {@link //Cell} class.</p>
//
// <p>WARNING: You don't call this class yourself. You access Sheet objects via the Book object that
// was returned when you called xlrd.open_workbook("myfile.xls").</p>

class Sheet
{
public:

    ////
    // Name of sheet.
    std::string name = "";

    ////
    // A reference to the Book object to which this sheet belongs.
    // Example usage: some_sheet.book.datemode
    SheetOwnerInterface* book = nullptr;
    
    ////
    // Number of rows in sheet. A row index is in range(thesheet.nrows).
    int nrows = 0;

    ////
    // Nominal number of columns in sheet. It is 1 + the maximum column index
    // found, ignoring trailing empty cells. See also open_workbook(ragged_rows=?)
    // and Sheet.{@link //Sheet.row_len}(row_index).
    int ncols = 0;

    ////
    // The map from a column index to a {@link //Colinfo} object. Often there is an entry
    // in COLINFO records for all column indexes in range(257).
    // Note that xlrd ignores the entry for the non-existent
    // 257th column. On the other hand, there may be no entry for unused columns.
    // <br /> -- New in version 0.6.1. Populated only if open_workbook(formatting_info=True).
    int colinfo_map;  // = {}

    ////
    // The map from a row index to a {@link //Rowinfo} object. Note that it is possible
    // to have missing entries -- at least one source of XLS files doesn't
    // bother writing ROW records.
    // <br /> -- New in version 0.6.1. Populated only if open_workbook(formatting_info=True).
    int rowinfo_map;  // = {}

    ////
    // List of address ranges of cells containing column labels.
    // These are set up in Excel by Insert > Name > Labels > Columns.
    // <br> -- New in version 0.6.0
    // <br>How to deconstruct the list:
    // <pre>
    // for crange in thesheet.col_label_ranges:
    //     rlo, rhi, clo, chi = crange
    //     for rx in xrange(rlo, rhi):
    //         for cx in xrange(clo, chi):
    //             print "Column label at (rowx=%d, colx=%d) is %r"
    //                 (rx, cx, thesheet.cell_value(rx, cx))
    // </pre>
    int col_label_ranges;  // = []

    ////
    // List of address ranges of cells containing row labels.
    // For more details, see <i>col_label_ranges</i> above.
    // <br> -- New in version 0.6.0
    int row_label_ranges; // = []

    ////
    // List of address ranges of cells which have been merged.
    // These are set up in Excel by Format > Cells > Alignment, then ticking
    // the "Merge cells" box.
    // <br> -- New in version 0.6.1. Extracted only if open_workbook(formatting_info=True).
    // <br>How to deconstruct the list:
    // <pre>
    // for crange in thesheet.merged_cells:
    //     rlo, rhi, clo, chi = crange
    //     for rowx in xrange(rlo, rhi):
    //         for colx in xrange(clo, chi):
    //             // cell (rlo, clo) (the top left one) will carry the data
    //             // and formatting info; the remainder will be recorded as
    //             // blank cells, but a renderer will apply the formatting info
    //             // for the top left cell (e.g. border, pattern) to all cells in
    //             // the range.
    // </pre>
    int merged_cells;  // = []
    
    ////
    // Mapping of (rowx, colx) to list of (offset, font_index) tuples. The offset
    // defines where in the string the font begins to be used.
    // Offsets are expected to be in ascending order.
    // If the first offset is not zero, the meaning is that the cell's XF's font should
    // be used from offset 0.
    // <br /> This is a sparse mapping. There is no entry for cells that are not formatted with  
    // rich text.
    // <br>How to use:
    // <pre>
    // runlist = thesheet.rich_text_runlist_map.get((rowx, colx))
    // if runlist:
    //     for offset, font_index in runlist:
    //         // do work here.
    //         pass
    // </pre>
    // Populated only if open_workbook(formatting_info=True).
    // <br /> -- New in version 0.7.2.
    // <br /> &nbsp;
    int rich_text_runlist_map;  // = {}    

    ////
    // Default column width from DEFCOLWIDTH record, else None.
    // From the OOo docs:<br />
    // """Column width in characters, using the width of the zero character
    // from default font (first FONT record in the file). Excel adds some
    // extra space to the default width, depending on the default font and
    // default font size. The algorithm how to exactly calculate the resulting
    // column width is not known.<br />
    // Example: The default width of 8 set in this record results in a column
    // width of 8.43 using Arial font with a size of 10 points."""<br />
    // For the default hierarchy, refer to the {@link //Colinfo} class.
    // <br /> -- New in version 0.6.1
    int defcolwidth;  // = None

    ////
    // Default column width from STANDARDWIDTH record, else None.
    // From the OOo docs:<br />
    // """Default width of the columns in 1/256 of the width of the zero
    // character, using default font (first FONT record in the file)."""<br />
    // For the default hierarchy, refer to the {@link //Colinfo} class.
    // <br /> -- New in version 0.6.1
    int standardwidth;  // = None

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the <i>optional</i> DEFAULTROWHEIGHT record.
    int default_row_height;  // = None

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the <i>optional</i> DEFAULTROWHEIGHT record.
    int default_row_height_mismatch;  // = None

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the <i>optional</i> DEFAULTROWHEIGHT record.
    int default_row_hidden; // = None

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the <i>optional</i> DEFAULTROWHEIGHT record.
    int default_additional_space_above;  // = None

    ////
    // Default value to be used for a row if there is
    // no ROW record for that row.
    // From the <i>optional</i> DEFAULTROWHEIGHT record.
    int default_additional_space_below;  // = None

    ////
    // Visibility of the sheet. 0 = visible, 1 = hidden (can be unhidden
    // by user -- Format/Sheet/Unhide), 2 = "very hidden" (can be unhidden
    // only by VBA macro).
    int visibility = 0;

    ////
    // A 256-element tuple corresponding to the contents of the GCW record for this sheet.
    // If no such record, treat as all bits zero.
    // Applies to BIFF4-7 only. See docs of the {@link //Colinfo} class for discussion.
    std::vector<int> gcw;  // = (0, ) * 256

    ////
    // <p>A list of {@link //Hyperlink} objects corresponding to HLINK records found
    // in the worksheet.<br />-- New in version 0.7.2 </p>
    int hyperlink_list;  // = []

    ////
    // <p>A sparse mapping from (rowx, colx) to an item in {@link //Sheet.hyperlink_list}.
    // Cells not covered by a hyperlink are not mapped.
    // It is possible using the Excel UI to set up a hyperlink that 
    // covers a larger-than-1x1 rectangle of cells.
    // Hyperlink rectangles may overlap (Excel doesn't check).
    // When a multiply-covered cell is clicked on, the hyperlink that is activated
    // (and the one that is mapped here) is the last in hyperlink_list.
    // <br />-- New in version 0.7.2 </p>
    int hyperlink_map;  // = {}

    ////
    // <p>A sparse mapping from (rowx, colx) to a {@link //Note} object.
    // Cells not containing a note ("comment") are not mapped.
    // <br />-- New in version 0.7.2 </p>
    int cell_note_map;  // = {}    
    
    ////
    // Number of columns in left pane (frozen panes; for split panes, see comments below in code)
    int vert_split_pos;  // = 0

    ////
    // Number of rows in top pane (frozen panes; for split panes, see comments below in code)
    int horz_split_pos = 0;

    ////
    // Index of first visible row in bottom frozen/split pane
    int horz_split_first_visible = 0;

    ////
    // Index of first visible column in right frozen/split pane
    int vert_split_first_visible = 0;

    ////
    // Frozen panes: ignore it. Split panes: explanation and diagrams in OOo docs.
    int split_active_pane = 0;

    ////
    // Boolean specifying if a PANE record was present, ignore unless you're xlutils.copy
    int has_pane_record = 0;

    ////
    // A list of the horizontal page breaks in this sheet.
    // Breaks are tuples in the form (index of row after break, start col index, end col index).
    // Populated only if open_workbook(formatting_info=True).
    // <br /> -- New in version 0.7.2
    std::vector<int> horizontal_page_breaks;  // = []

    ////
    // A list of the vertical page breaks in this sheet.
    // Breaks are tuples in the form (index of col after break, start row index, end row index).
    // Populated only if open_workbook(formatting_info=True).
    // <br /> -- New in version 0.7.2
    std::vector<int> vertical_page_breaks;  // = []

    int biff_version;
    int _position;
    int number;
    int verbosity;
    int formatting_info;
    int ragged_rows;
    std::map<int, int>* _xf_index_to_xl_type_map;
    int _maxdatarowx = -1; // highest rowx containing a non-empty cell
    int _maxdatacolx = -1; // highest colx containing a non-empty cell
    int _dimnrows = 0; // as per DIMENSIONS record
    int _dimncols = 0;
    std::vector<utils::any> _cell_values;
    std::vector<utils::any> _cell_types;
    std::vector<int> _cell_xf_indexes;
    std::vector<int> _xf_index_stats;

    // _WINDOW2_options
    int show_formulas;
    int show_grid_lines;
    int show_sheet_headers;
    int panes_are_frozen;
    int show_zero_values;
    int automatic_grid_line_colour;
    int columns_from_right_to_left;
    int show_outline_symbols;
    int remove_splits_if_pane_freeze_is_removed;
    // Multiple sheets can be selected, but only one can be active
    // {hold down Ctrl and click multiple tabs in the file in OOo}
    int sheet_selected;
    // "sheet_visible" should really be called "sheet_active"
    // and is 1 when this sheet is the sheet displayed when the file
    // is open. More than likely only one sheet should ever be set as
    // visible.
    // This would correspond to the Book's sheet_active attribute, but
    // that doesn't exist as WINDOW1 records aren't currently processed.
    // The real thing is the visibility attribute from the BOUNDSHEET record.
    int sheet_visible;
    int show_in_page_break_preview;

    int first_visible_rowx;
    int first_visible_colx;
    int gridline_colour_index;
    int gridline_colour_rgb;

    int cooked_page_break_preview_mag_factor;
    int cooked_normal_view_mag_factor;
    int cached_page_break_preview_mag_factor = 0;  // default (60%), from WINDOW2 record
    int cached_normal_view_mag_factor = 0;  // default (100%), from WINDOW2 record
    int scl_mag_factor = 0; // from SCL record
    int _ixfe = 0; // BIFF2 only
    int _cell_attr_to_xfx; // BIFF2.0 only
    int utter_max_rows;
    int utter_max_cols;
    int _first_full_rowx;

    inline
    Sheet(SheetOwnerInterface& owner, int position, std::string name, int number)
    {
        this->book = &owner;
        this->biff_version = owner.biff_version;
        this->_position = position;
        // this->logfile = book.logfile;
        // this->bt = array('B', [XL_CELL_EMPTY])
        // this->bf = array('h', [-1])
        this->name = name;
        this->number = number;
        this->verbosity = owner.verbosity;
        this->formatting_info = owner.formatting_info;
        this->ragged_rows = owner.ragged_rows;

        this->_xf_index_to_xl_type_map = owner._xf_index_to_xl_type_map;
        this->nrows = 0; // actual, including possibly empty cells
        this->ncols = 0;
        this->_maxdatarowx = -1; // highest rowx containing a non-empty cell
        this->_maxdatacolx = -1; // highest colx containing a non-empty cell
        this->_dimnrows = 0; // as per DIMENSIONS record
        this->_dimncols = 0;
        this->_cell_values = {};
        this->_cell_types = {};
        this->_cell_xf_indexes = {};
        this->defcolwidth = 0;
        this->standardwidth = 0;
        this->default_row_height = 0;
        this->default_row_height_mismatch = 0;
        this->default_row_hidden = 0;
        this->default_additional_space_above = 0;
        this->default_additional_space_below = 0;
        this->colinfo_map = {};
        this->rowinfo_map = {};
        this->col_label_ranges = {};
        this->row_label_ranges = {};
        this->merged_cells = {};
        this->rich_text_runlist_map = {};
        this->horizontal_page_breaks = {};
        this->vertical_page_breaks = {};
        this->_xf_index_stats[0] = 0;
        this->_xf_index_stats[1] = 1;
        this->_xf_index_stats[2] = 2;
        this->_xf_index_stats[3] = 3;
        this->visibility = owner._sheet_visibility[number]; // from BOUNDSHEET record
        // for attr, defval in _WINDOW2_options:
        //     setattr(self, attr, defval)
        {
            // _WINDOW2_options
            this->show_formulas = 0;
            this->show_grid_lines = 1;
            this->show_sheet_headers = 1;
            this->panes_are_frozen = 0;
            this->show_zero_values = 1;
            this->automatic_grid_line_colour = 1;
            this->columns_from_right_to_left = 0;
            this->show_outline_symbols = 1;
            this->remove_splits_if_pane_freeze_is_removed = 0;
            // Multiple sheets can be selected, but only one can be active
            // {hold down Ctrl and click multiple tabs in the file in OOo}
            this->sheet_selected = 0;
            // "sheet_visible" should really be called "sheet_active"
            // and is 1 when this sheet is the sheet displayed when the file
            // is open. More than likely only one sheet should ever be set as
            // visible.
            // This would correspond to the Book's sheet_active attribute, but
            // that doesn't exist as WINDOW1 records aren't currently processed.
            // The real thing is the visibility attribute from the BOUNDSHEET record.
            this->sheet_visible = 0;
            this->show_in_page_break_preview = 0;
        }
        this->first_visible_rowx = 0;
        this->first_visible_colx = 0;
        this->gridline_colour_index = 0x40;
        this->gridline_colour_rgb = 0; // pre-BIFF8
        this->hyperlink_list = {};
        this->hyperlink_map = {};
        this->cell_note_map = {};

        // Values calculated by xlrd to predict the mag factors that
        // will actually be used by Excel to display your worksheet.
        // Pass these values to xlwt when writing XLS files.
        // Warning 1: Behaviour of OOo Calc and Gnumeric has been observed to differ from Excel's.
        // Warning 2: A value of zero means almost exactly what it says. Your sheet will be
        // displayed as a very tiny speck on the screen. xlwt will reject attempts to set
        // a mag_factor that is not (10 <= mag_factor <= 400).
        this->cooked_page_break_preview_mag_factor = 60;
        this->cooked_normal_view_mag_factor = 100;

        // Values (if any) actually stored on the XLS file
        this->cached_page_break_preview_mag_factor = 0;  // default (60%), from WINDOW2 record
        this->cached_normal_view_mag_factor = 0;  // default (100%), from WINDOW2 record
        this->scl_mag_factor = 0; // from SCL record

        this->_ixfe = 0; // BIFF2 only
        this->_cell_attr_to_xfx = {}; // BIFF2.0 only

        //////// Don't initialise this here, use class attribute initialisation.
        //////// this->gcw = (0, ) * 256 ////////

        if (this->biff_version >= 80) {
            this->utter_max_rows = 65536;
        }
        else {
            this->utter_max_rows = 16384;
        }
        this->utter_max_cols = 256;

        this->_first_full_rowx = -1;

    };

    ////
    // {@link //Cell} object in the given row and column.
    Cell cell(int rowx, int colx);

    ////
    // Value of the cell in the given row and column.
    utils::any cell_value(int rowx, int colx);

    ////
    // Type of the cell in the given row and column.
    // Refer to the documentation of the {@link //Cell} class.
    int cell_type(int rowx, int colx);

    ////
    // XF index of the cell in the given row and column.
    // This is an index into Book.{@link //Book.xf_list}.
    // <br /> -- New in version 0.6.1
    int cell_xf_index(int rowx, int colx);

    ////
    // Returns the effective number of cells in the given row. For use with
    // open_workbook(ragged_rows=True) which is likely to produce rows
    // with fewer than {@link //Sheet.ncols} cells.
    // <br /> -- New in version 0.7.2
    int row_len(int rowx);

    ////
    // Returns a sequence of the {@link //Cell} objects in the given row.
    std::vector<Cell> row(int rowx);

    ////
    // Returns a generator for iterating through each row.
    std::vector<std::vector<Cell>> get_rows();

    ////
    // Returns a slice of the types
    // of the cells in the given row.
    std::vector<int> row_types(int rowx, int start_colx=0, int end_colx=-1);

    ////
    // Returns a slice of the values
    // of the cells in the given row.
    std::vector<utils::any> row_values(int rowx, int start_colx=0, int end_colx=-1);

    ////
    // Returns a slice of the {@link //Cell} objects in the given row.
    std::vector<Cell> row_slice(int rowx, int start_colx=0, int end_colx=-1);

    ////
    // Returns a slice of the {@link //Cell} objects in the given column.
    std::vector<Cell> col_slice(int colx, int start_rowx=0, int end_rowx=-1);

    ////
    // Returns a slice of the values of the cells in the given column.
    std::vector<utils::any> col_values(int colx, int start_rowx=0, int end_rowx=-1);

    ////
    // Returns a slice of the types of the cells in the given column.
    std::vector<int> col_types(int colx, int start_rowx=0, int end_rowx=-1);

    ////
    // Returns a sequence of the {@link //Cell} objects in the given column.
    std::vector<Cell> col(int colx);

    // === Following methods are used in building the worksheet.
    // === They are not part of the API.

    void tidy_dimensions();

    inline
    void put_cell(int rowx, int colx, int ctype, utils::any value, int xf_index) {
        if (this->ragged_rows) {
            this->put_cell_ragged(rowx, colx, ctype, value, xf_index);
        }
        else {
            this->put_cell_unragged(rowx, colx, ctype, value, xf_index);
        }
    };

    inline
    void put_cell_ragged(int rowx, int colx, int ctype, utils::any value, int xf_index)
    {

    }

    inline
    void put_cell_unragged(int rowx, int colx, int ctype, utils::any value, int xf_index)
    {
        
    }

    // === Methods after this line neither know nor care about how cells are stored.

    int read(void* bk);

    void string_record_contents(std::vector<uint8_t> data);

    void update_cooked_mag_factors();

    void fixed_BIFF2_xfindex(int cell_attr, int rowx, int colx, int true_xfx=0);

    void insert_new_BIFF20_xf(int cell_attr, int style=0);

    void fake_XF_from_BIFF20_cell_attr(int cell_attr, int style=0);

    void req_fmt_info();

    ////
    // Determine column display width.
    // <br /> -- New in version 0.6.1
    // <br />
    // @param colx Index of the queried column, range 0 to 255.
    // Note that it is possible to find out the width that will be used to display
    // columns with no cell information e.g. column IV (colx=255).
    // @return The column width that will be used for displaying
    // the given column by Excel, in units of 1/256th of the width of a
    // standard character (the digit zero in the first font).

    inline
    void computed_column_width(int colx);

    inline
    void handle_hlink(std::vector<uint8_t> data);

    inline
    void handle_quicktip(std::vector<uint8_t> data);

    inline
    void handle_msodrawingetc(int recid, int data_len,
                              std::vector<uint8_t> data);

    inline
    void handle_obj(std::vector<uint8_t> data);

    inline
    void handle_note(u8vec data, int txos);

    inline
    void handle_txo(u8vec data);

    inline
    void handle_feat11(u8vec data);
};

class MSODrawing {};

class MSObj {};

class MSTxo {};

////    
// <p> Represents a user "comment" or "note".
// Note objects are accessible through Sheet.{@link //Sheet.cell_note_map}.
// <br />-- New in version 0.7.2  
// </p>
class Note
{
public:
    ////
    // Author of note
    std::string author = "";
    ////
    // True if the containing column is hidden
    int col_hidden = 0;
    ////
    // Column index
    int colx = 0;
    ////
    // List of (offset_in_string, font_index) tuples.
    // Unlike Sheet.{@link //Sheet.rich_text_runlist_map}, the first offset should always be 0.
    std::vector<std::tuple<int, int>> rich_text_runlist;
    ////
    // True if the containing row is hidden
    int row_hidden = 0;
    ////
    // Row index
    int rowx = 0;
    ////
    // True if note is always shown
    int show = 0;
    ////
    // Text of the note
    std::string text;
};

////
// <p>Contains the attributes of a hyperlink.
// Hyperlink objects are accessible through Sheet.{@link //Sheet.hyperlink_list}
// and Sheet.{@link //Sheet.hyperlink_map}.
// <br />-- New in version 0.7.2
// </p>   
class Hyperlink
{
public:
    ////
    // Index of first row
    int frowx = -1;
    ////
    // Index of last row
    int lrowx = -1;
    ////
    // Index of first column
    int fcolx = -1;
    ////
    // Index of last column
    int lcolx = -1;
    ////
    // Type of hyperlink. Unicode string, one of 'url', 'unc',
    // 'local file', 'workbook', 'unknown'
    std::string type;
    ////
    // The URL or file-path, depending in the type. Unicode string, except 
    // in the rare case of a local but non-existent file with non-ASCII
    // characters in the name, in which case only the "8.3" filename is available,
    // as a bytes (3.x) or str (2.x) string, <i>with unknown encoding.</i>
    std::string url_or_path;
    ////
    // Description ... this is displayed in the cell,
    // and should be identical to the cell value. Unicode string, or None. It seems
    // impossible NOT to have a description created by the Excel UI.
    std::string desc;
    ////
    // Target frame. Unicode string. Note: I have not seen a case of this.
    // It seems impossible to create one in the Excel UI.
    std::string target;
    ////
    // "Textmark": the piece after the "//" in 
    // "http://docs.python.org/library//struct_module", or the Sheet1!A1:Z99
    // part when type is "workbook".
    std::string textmark;
    ////
    // The text of the "quick tip" displayed when the cursor
    // hovers over the hyperlink.
    std::string quicktip;
};

// === helpers ===

double unpack_RK(u8vec rk_str) {
    uint8_t flags = rk_str[0];
    if (flags & 2) {
        // There's a SIGNED 30-bit integer in there!
        int i  = utils::as_int32(rk_str);
        i >>= 2; // div by 4 to drop the 2 flag bits
        if (flags & 1) {
            return i / 100.0;
        }
        return double(i);
    }
    else {
        // It's the most significant 30 bits of an IEEE 754 64-bit FP number
        u8vec buf = {0, 0, 0, 0};
        buf.push_back(flags & 252);
        buf.push_back(rk_str[1]);
        buf.push_back(rk_str[2]);
        buf.push_back(rk_str[3]);
        double d = utils::as_double(buf);
        if (flags & 1) {
            return d / 100.0;
        }
        return d;
    }
}

////////// =============== Cell ======================================== //////////
const std::map<int, int>
cellty_from_fmtty = {
    {FNU, XL_CELL_NUMBER},
    {FUN, XL_CELL_NUMBER},
    {FGE, XL_CELL_NUMBER},
    {FDT, XL_CELL_DATE},
    {FTX, XL_CELL_NUMBER}, // Yes, a number can be formatted as text.
};

const std::map<int, const std::string>
ctype_text = {
    {XL_CELL_EMPTY, "empty"},
    {XL_CELL_TEXT, "text"},
    {XL_CELL_NUMBER, "number"},
    {XL_CELL_DATE, "xldate"},
    {XL_CELL_BOOLEAN, "bool"},
    {XL_CELL_ERROR, "error"},
    {XL_CELL_BLANK, "blank"},
};

////
// <p>Contains the data for one cell.</p>
//
// <p>WARNING: You don't call this class yourself. You access Cell objects
// via methods of the {@link //Sheet} object(s) that you found in the {@link //Book} object that
// was returned when you called xlrd.open_workbook("myfile.xls").</p>
// <p> Cell objects have three attributes: <i>ctype</i> is an int, <i>value</i>
// (which depends on <i>ctype</i>) and <i>xf_index</i>.
// If "formatting_info" is not enabled when the workbook is opened, xf_index will be None.
// The following table describes the types of cells and how their values
// are represented in Python.</p>
//
// <table border="1" cellpadding="7">
// <tr>
// <th>Type symbol</th>
// <th>Type number</th>
// <th>Python value</th>
// </tr>
// <tr>
// <td>XL_CELL_EMPTY</td>
// <td align="center">0</td>
// <td>empty string u''</td>
// </tr>
// <tr>
// <td>XL_CELL_TEXT</td>
// <td align="center">1</td>
// <td>a Unicode string</td>
// </tr>
// <tr>
// <td>XL_CELL_NUMBER</td>
// <td align="center">2</td>
// <td>float</td>
// </tr>
// <tr>
// <td>XL_CELL_DATE</td>
// <td align="center">3</td>
// <td>float</td>
// </tr>
// <tr>
// <td>XL_CELL_BOOLEAN</td>
// <td align="center">4</td>
// <td>int; 1 means TRUE, 0 means FALSE</td>
// </tr>
// <tr>
// <td>XL_CELL_ERROR</td>
// <td align="center">5</td>
// <td>int representing internal Excel codes; for a text representation,
// refer to the supplied dictionary error_text_from_code</td>
// </tr>
// <tr>
// <td>XL_CELL_BLANK</td>
// <td align="center">6</td>
// <td>empty string u''. Note: this type will appear only when
// open_workbook(..., formatting_info=True) is used.</td>
// </tr>
// </table>
//<p></p>

class Cell 
{
public:
    int ctype;
    utils::any value;
    int xf_index;

    inline
    Cell(int ctype, utils::any value, int xf_index=-1) {
        this->ctype = ctype;
        this->value = value;
        this->xf_index = xf_index;
    }

    // def __repr__(self):
    //     if self.xf_index is None:
    //         return "%s:%r" % (ctype_text[self.ctype], self.value)
    //     else:
    //         return "%s:%r (XF:%r)" % (ctype_text[self.ctype], self.value, self.xf_index)
};

Cell empty_cell = Cell(biffh::XL_CELL_EMPTY, "");

////////// =============== Colinfo and Rowinfo ============================== //////////

////
// Width and default formatting information that applies to one or
// more columns in a sheet. Derived from COLINFO records.
//
// <p> Here is the default hierarchy for width, according to the OOo docs:
//
// <br />"""In BIFF3, if a COLINFO record is missing for a column,
// the width specified in the record DEFCOLWIDTH is used instead.
//
// <br />In BIFF4-BIFF7, the width set in this [COLINFO] record is only used,
// if the corresponding bit for this column is cleared in the GCW
// record, otherwise the column width set in the DEFCOLWIDTH record
// is used (the STANDARDWIDTH record is always ignored in this case [see footnote!]).
//
// <br />In BIFF8, if a COLINFO record is missing for a column,
// the width specified in the record STANDARDWIDTH is used.
// If this [STANDARDWIDTH] record is also missing,
// the column width of the record DEFCOLWIDTH is used instead."""
// <br />
//
// Footnote:  The docs on the GCW record say this:
// """<br />
// If a bit is set, the corresponding column uses the width set in the STANDARDWIDTH
// record. If a bit is cleared, the corresponding column uses the width set in the
// COLINFO record for this column.
// <br />If a bit is set, and the worksheet does not contain the STANDARDWIDTH record, or if
// the bit is cleared, and the worksheet does not contain the COLINFO record, the DEFCOLWIDTH
// record of the worksheet will be used instead.
// <br />"""<br />
// At the moment (2007-01-17) xlrd is going with the GCW version of the story.
// Reference to the source may be useful: see the computed_column_width(colx) method
// of the Sheet class.
// <br />-- New in version 0.6.1
// </p>

class Colinfo
{
public:
    ////
    // Width of the column in 1/256 of the width of the zero character,
    // using default font (first FONT record in the file).
    int width = 0;
    ////
    // XF index to be used for formatting empty cells.
    int xf_index = -1;
    ////
    // 1 = column is hidden
    int hidden = 0;
    ////
    // Value of a 1-bit flag whose purpose is unknown
    // but is often seen set to 1
    int bit1_flag = 0;
    ////
    // Outline level of the column, in range(7).
    // (0 = no outline)
    int outline_level = 0;
    ////
    // 1 = column is collapsed
    int collapsed = 0;
};

const int _USE_SLOTS = 1;

////
// <p>Height and default formatting information that applies to a row in a sheet.
// Derived from ROW records.
// <br /> -- New in version 0.6.1</p>
//
// <p><b>height</b>: Height of the row, in twips. One twip == 1/20 of a point.</p>
//
// <p><b>has_default_height</b>: 0 = Row has custom height; 1 = Row has default height.</p>
//
// <p><b>outline_level</b>: Outline level of the row (0 to 7) </p>
//
// <p><b>outline_group_starts_ends</b>: 1 = Outline group starts or ends here (depending on where the
// outline buttons are located, see WSBOOL record [TODO ??]),
// <i>and</i> is collapsed </p>
//
// <p><b>hidden</b>: 1 = Row is hidden (manually, or by a filter or outline group) </p>
//
// <p><b>height_mismatch</b>: 1 = Row height and default font height do not match </p>
//
// <p><b>has_default_xf_index</b>: 1 = the xf_index attribute is usable; 0 = ignore it </p>
//
// <p><b>xf_index</b>: Index to default XF record for empty cells in this row.
// Don't use this if has_default_xf_index == 0. </p>
//
// <p><b>additional_space_above</b>: This flag is set, if the upper border of at least one cell in this row
// or if the lower border of at least one cell in the row above is
// formatted with a thick line style. Thin and medium line styles are not
// taken into account. </p>
//
// <p><b>additional_space_below</b>: This flag is set, if the lower border of at least one cell in this row
// or if the upper border of at least one cell in the row below is
// formatted with a medium or thick line style. Thin line styles are not
// taken into account. </p>

class Rowinfo
{
public:
    int height;
    int has_default_height;
    int outline_level;
    int outline_group_starts_ends;
    int hidden;
    int height_mismatch;
    int has_default_xf_index;
    int xf_index;
    int additional_space_above;
    int additional_space_below;

};


}
}
