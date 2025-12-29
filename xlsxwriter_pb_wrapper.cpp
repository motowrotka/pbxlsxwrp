#include <windows.h>
#include "xlsxwriter.h"
#include <string>

// PowerBuilder 9.0 compatible wrapper for libxlsxwriter
// Uses __stdcall convention and C-style strings

extern "C" {

// Workbook handle type
typedef void* PB_WORKBOOK;
typedef void* PB_WORKSHEET;
typedef void* PB_FORMAT;

// Create a new workbook
__declspec(dllexport) PB_WORKBOOK __stdcall pb_workbook_new(const char* filename) {
    if (!filename) return NULL;
    setlocale(LC_NUMERIC, "C");
    return (PB_WORKBOOK)workbook_new(filename);
}

// Add a worksheet
__declspec(dllexport) PB_WORKSHEET __stdcall pb_worksheet_add(PB_WORKBOOK workbook, const char* sheetname) {
    if (!workbook) return NULL;
    return (PB_WORKSHEET)workbook_add_worksheet((lxw_workbook*)workbook, sheetname);
}

// Write string to cell
__declspec(dllexport) int __stdcall pb_worksheet_write_string(
    PB_WORKSHEET worksheet, 
    int row, 
    int col, 
    const char* text,
    PB_FORMAT format) {
    
    if (!worksheet || !text) return -1;
    return worksheet_write_string((lxw_worksheet*)worksheet, row, col, text, (lxw_format*)format);
}

// Write number to cell
__declspec(dllexport) int __stdcall pb_worksheet_write_number(
    PB_WORKSHEET worksheet,
    int row,
    int col,
    double number,
    PB_FORMAT format) {
    
    if (!worksheet) return -1;
    return worksheet_write_number((lxw_worksheet*)worksheet, row, col, number, (lxw_format*)format);
}

// Write formula to cell
__declspec(dllexport) int __stdcall pb_worksheet_write_formula(
    PB_WORKSHEET worksheet,
    int row,
    int col,
    const char* formula,
    PB_FORMAT format) {
    
    if (!worksheet || !formula) return -1;
    return worksheet_write_formula((lxw_worksheet*)worksheet, row, col, formula, (lxw_format*)format);
}

// Write datetime to cell
__declspec(dllexport) int __stdcall pb_worksheet_write_datetime(
    PB_WORKSHEET worksheet,
    int row,
    int col,
    int year,
    int month,
    int day,
    int hour,
    int min,
    double sec,
    PB_FORMAT format) {
    
    if (!worksheet) return -1;
    
    lxw_datetime dt;
    dt.year = year;
    dt.month = month;
    dt.day = day;
    dt.hour = hour;
    dt.min = min;
    dt.sec = sec;
    
    return worksheet_write_datetime((lxw_worksheet*)worksheet, row, col, &dt, (lxw_format*)format);
}

// Set column width
__declspec(dllexport)
int __stdcall pb_worksheet_set_column(
    PB_WORKSHEET worksheet,
    int first_col,
    int last_col,
    double width,
    PB_FORMAT format)
{
    if (!worksheet) return -1;

    lxw_format* fmt = NULL;
    if (format != NULL && format != 0)
        fmt = (lxw_format*)format;
    
    return worksheet_set_column((lxw_worksheet*)worksheet, first_col, last_col, width, fmt);
}

__declspec(dllexport)
int __stdcall pb_worksheet_set_row(
    PB_WORKSHEET worksheet,
    int row,
    double height,
    PB_FORMAT format)
{
    if (!worksheet) return -1;

    lxw_format* fmt = NULL;
    if (format != NULL && format != 0)
        fmt = (lxw_format*)format;

    return worksheet_set_row((lxw_worksheet*)worksheet, row, height, fmt);
}

// Add format
__declspec(dllexport) PB_FORMAT __stdcall pb_workbook_add_format(PB_WORKBOOK workbook) {
    if (!workbook) return NULL;
    return (PB_FORMAT)workbook_add_format((lxw_workbook*)workbook);
}

// Format functions
__declspec(dllexport) void __stdcall pb_format_set_bold(PB_FORMAT format) {
    if (format) format_set_bold((lxw_format*)format);
}

__declspec(dllexport) void __stdcall pb_format_set_italic(PB_FORMAT format) {
    if (format) format_set_italic((lxw_format*)format);
}

__declspec(dllexport) void __stdcall pb_format_set_font_size(PB_FORMAT format, int size) {
    if (format) format_set_font_size((lxw_format*)format, size);
}

__declspec(dllexport) void __stdcall pb_format_set_font_color(PB_FORMAT format, int color) {
    if (format) format_set_font_color((lxw_format*)format, color);
}

__declspec(dllexport) void __stdcall pb_format_set_bg_color(PB_FORMAT format, int color) {
    if (format) format_set_bg_color((lxw_format*)format, color);
}

__declspec(dllexport) void __stdcall pb_format_set_align(PB_FORMAT format, int align) {
    if (format) format_set_align((lxw_format*)format, align);
}

__declspec(dllexport) void __stdcall pb_format_set_num_format(PB_FORMAT format, const char* num_format) {
    if (format && num_format) format_set_num_format((lxw_format*)format, num_format);
}

__declspec(dllexport) void __stdcall pb_format_set_border(PB_FORMAT format, int style) {
    if (format) format_set_border((lxw_format*)format, style);
}

// Merge cells
__declspec(dllexport) int __stdcall pb_worksheet_merge_range(
    PB_WORKSHEET worksheet,
    int first_row,
    int first_col,
    int last_row,
    int last_col,
    const char* text,
    PB_FORMAT format) {
    
    if (!worksheet || !text) return -1;
    return worksheet_merge_range((lxw_worksheet*)worksheet, first_row, first_col, last_row, last_col, text, (lxw_format*)format);
}

// Insert image
__declspec(dllexport) int __stdcall pb_worksheet_insert_image(
    PB_WORKSHEET worksheet,
    int row,
    int col,
    const char* filename) {
    
    if (!worksheet || !filename) return -1;
    return worksheet_insert_image((lxw_worksheet*)worksheet, row, col, filename);
}

// Autofilter
__declspec(dllexport) int __stdcall pb_worksheet_autofilter(
    PB_WORKSHEET worksheet,
    int first_row,
    int first_col,
    int last_row,
    int last_col) {
    
    if (!worksheet) return -1;
    return worksheet_autofilter((lxw_worksheet*)worksheet, first_row, first_col, last_row, last_col);
}

__declspec(dllexport)
int __stdcall pb_worksheet_autofit_column(
    PB_WORKSHEET ws,
    int col,
    int max_chars,
    PB_FORMAT format)
{
    lxw_worksheet* w = safe_ws(ws);
    if (!w) return -1;

    if (max_chars <= 0) {
        // fallback – np. standardowa szerokość 8.43
        return worksheet_set_column(w, col, col, 8.43, safe_fmt(format));
    }

    // Heurystyka: Excelowa szerokość ~ liczba znaków * współczynnik
    // Dla Calibri 11 mniej więcej 1:1, ale dodamy trochę zapasu.
    double padding = 1.0;      // dodatkowy zapas na margines/padding
    double width = max_chars + padding;

    return worksheet_set_column(w, col, col, width, safe_fmt(format));
}


// Close workbook
__declspec(dllexport) int __stdcall pb_workbook_close(PB_WORKBOOK workbook) {
    if (!workbook) return -1;
    return workbook_close((lxw_workbook*)workbook);
}

// Get version info
__declspec(dllexport) const char* __stdcall pb_get_version() {
    return "libxlsxwriter PowerBuilder Wrapper v1.0";
}

} // extern "C"
