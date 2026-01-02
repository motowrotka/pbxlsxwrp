#include <windows.h>
#include <string>
#include <locale.h>
#include <cstdint>
#include "xlsxwriter.h"

extern "C" {

typedef void* PB_WORKBOOK;
typedef void* PB_WORKSHEET;
typedef void* PB_FORMAT;


// ------------------------------------------------------------
//  ANSI → UTF‑8 conversion (PB9 → Excel-safe)
// ------------------------------------------------------------
static std::string ansi_to_utf8(const char* ansi)
{
    if (!ansi)
        return std::string();

    // ANSI → UTF‑16
    int wide_len = MultiByteToWideChar(CP_ACP, 0, ansi, -1, NULL, 0);
    if (wide_len <= 0)
        return std::string();

    std::wstring wstr;
    wstr.resize(wide_len);
    MultiByteToWideChar(CP_ACP, 0, ansi, -1, &wstr[0], wide_len);

    // UTF‑16 → UTF‑8
    int utf8_len = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, NULL, 0, NULL, NULL);
    if (utf8_len <= 0)
        return std::string();

    std::string utf8;
    utf8.resize(utf8_len);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &utf8[0], utf8_len, NULL, NULL);

    return utf8;
}


// ------------------------------------------------------------
//  Helpers
// ------------------------------------------------------------
static inline lxw_worksheet* safe_ws(PB_WORKSHEET ws) {
    return (lxw_worksheet*)ws;
}

static inline lxw_format* safe_fmt(PB_FORMAT fmt) {
    return (lxw_format*)fmt;
}


// ------------------------------------------------------------
//  Workbook creation
// ------------------------------------------------------------
__declspec(dllexport) PB_WORKBOOK __stdcall pb_workbook_new(const char* filename) {
    if (!filename) return NULL;

    setlocale(LC_NUMERIC, "C");

    std::string utf8_filename = ansi_to_utf8(filename);
    if (utf8_filename.empty())
        return NULL;

    //return (PB_WORKBOOK)workbook_new(utf8_filename.c_str());
    return (PB_WORKBOOK)workbook_new(filename.c_str());
}


// ------------------------------------------------------------
//  Worksheet creation
// ------------------------------------------------------------
__declspec(dllexport) PB_WORKSHEET __stdcall pb_worksheet_add(PB_WORKBOOK workbook, const char* sheetname) {
    if (!workbook) return NULL;

    const char* name_ptr = NULL;
    std::string utf8_name;

    if (sheetname && sheetname[0] != '\0') {
        utf8_name = ansi_to_utf8(sheetname);
        name_ptr = utf8_name.c_str();
    }

    return (PB_WORKSHEET)workbook_add_worksheet((lxw_workbook*)workbook, name_ptr);
}


// ------------------------------------------------------------
//  Write string
// ------------------------------------------------------------
__declspec(dllexport) int __stdcall pb_worksheet_write_string(
    PB_WORKSHEET worksheet,
    std::int32_t row,
    std::int32_t col,
    const char* text,
    PB_FORMAT format)
{
    if (!worksheet || !text) return -1;

    std::string utf8_text = ansi_to_utf8(text);

    return worksheet_write_string(
        safe_ws(worksheet),
        row,
        col,
        utf8_text.c_str(),
        safe_fmt(format)
    );
}


// ------------------------------------------------------------
//  Write number
// ------------------------------------------------------------
__declspec(dllexport) int __stdcall pb_worksheet_write_number(
    PB_WORKSHEET worksheet,
    std::int32_t row,
    std::int32_t col,
    double number,
    PB_FORMAT format)
{
    if (!worksheet) return -1;
    return worksheet_write_number(safe_ws(worksheet), row, col, number, safe_fmt(format));
}


// ------------------------------------------------------------
//  Write formula
// ------------------------------------------------------------
__declspec(dllexport) int __stdcall pb_worksheet_write_formula(
    PB_WORKSHEET worksheet,
    std::int32_t row,
    std::int32_t col,
    const char* formula,
    PB_FORMAT format)
{
    if (!worksheet || !formula) return -1;

    std::string utf8_formula = ansi_to_utf8(formula);

    return worksheet_write_formula(
        safe_ws(worksheet),
        row,
        col,
        utf8_formula.c_str(),
        safe_fmt(format)
    );
}


// ------------------------------------------------------------
//  Write datetime
// ------------------------------------------------------------
__declspec(dllexport) int __stdcall pb_worksheet_write_datetime(
    PB_WORKSHEET worksheet,
    std::int32_t row,
    std::int32_t col,
    int year,
    int month,
    int day,
    int hour,
    int min,
    double sec,
    PB_FORMAT format)
{
    if (!worksheet) return -1;

    lxw_datetime dt;
    dt.year  = year;
    dt.month = month;
    dt.day   = day;
    dt.hour  = hour;
    dt.min   = min;
    dt.sec   = sec;

    return worksheet_write_datetime(safe_ws(worksheet), row, col, &dt, safe_fmt(format));
}

// ------------------------------------------------------------
//  Write full row (fast path)
// ------------------------------------------------------------
__declspec(dllexport) int __stdcall pb_worksheet_write_row(
    PB_WORKSHEET worksheet,
    std::int32_t row,
    const char** values,
    const int* visible,
    PB_FORMAT* formats,
    std::int32_t colcount)
{
    if (!worksheet || !values || !visible) 
        return -1;

    lxw_worksheet* ws = safe_ws(worksheet);

    int excel_col = 0;

    for (int i = 0; i < colcount; i++) {

        if (visible[i]) {

            const char* v = values[i];
            if (!v) v = "";

            worksheet_write_string(
                ws,
                row,
                excel_col,
                v,
                safe_fmt(formats ? formats[i] : NULL)
            );

            excel_col++;
        }
    }

    return 0;
}

// ------------------------------------------------------------
//  Column / row formatting
// ------------------------------------------------------------
__declspec(dllexport)
int __stdcall pb_worksheet_set_column(
    PB_WORKSHEET worksheet,
    std::int32_t first_col,
    std::int32_t last_col,
    double width,
    PB_FORMAT format)
{
    if (!worksheet) return -1;
    return worksheet_set_column(safe_ws(worksheet), first_col, last_col, width, safe_fmt(format));
}

__declspec(dllexport)
int __stdcall pb_worksheet_set_row(
    PB_WORKSHEET worksheet,
    std::int32_t row,
    double height,
    PB_FORMAT format)
{
    if (!worksheet) return -1;
    return worksheet_set_row(safe_ws(worksheet), row, height, safe_fmt(format));
}


// ------------------------------------------------------------
//  Format creation
// ------------------------------------------------------------
__declspec(dllexport) PB_FORMAT __stdcall pb_workbook_add_format(PB_WORKBOOK workbook) {
    if (!workbook) return NULL;
    return (PB_FORMAT)workbook_add_format((lxw_workbook*)workbook);
}


// ------------------------------------------------------------
//  Format setters
// ------------------------------------------------------------
__declspec(dllexport) void __stdcall pb_format_set_bold(PB_FORMAT format) {
    if (format) format_set_bold(safe_fmt(format));
}

__declspec(dllexport) void __stdcall pb_format_set_italic(PB_FORMAT format) {
    if (format) format_set_italic(safe_fmt(format));
}

__declspec(dllexport) void __stdcall pb_format_set_font_size(PB_FORMAT format, int size) {
    if (format) format_set_font_size(safe_fmt(format), size);
}

__declspec(dllexport) void __stdcall pb_format_set_font_color(PB_FORMAT format, int color) {
    if (format) format_set_font_color(safe_fmt(format), color);
}

__declspec(dllexport) void __stdcall pb_format_set_bg_color(PB_FORMAT format, int color) {
    if (format) format_set_bg_color(safe_fmt(format), color);
}

__declspec(dllexport) void __stdcall pb_format_set_align(PB_FORMAT format, int align) {
    if (format) format_set_align(safe_fmt(format), align);
}

__declspec(dllexport) void __stdcall pb_format_set_num_format(PB_FORMAT format, const char* num_format) {
    if (format && num_format) {
        std::string utf8 = ansi_to_utf8(num_format);
        format_set_num_format(safe_fmt(format), utf8.c_str());
    }
}

__declspec(dllexport) void __stdcall pb_format_set_border(PB_FORMAT format, int style) {
    if (format) format_set_border(safe_fmt(format), style);
}


// ------------------------------------------------------------
//  NEW: Enable wrap text
// ------------------------------------------------------------
__declspec(dllexport) void __stdcall pb_format_set_text_wrap(PB_FORMAT format) {
    if (format) format_set_text_wrap(safe_fmt(format));
}


// ------------------------------------------------------------
//  Merge cells
// ------------------------------------------------------------
__declspec(dllexport) int __stdcall pb_worksheet_merge_range(
    PB_WORKSHEET worksheet,
    std::int32_t first_row,
    std::int32_t first_col,
    std::int32_t last_row,
    std::int32_t last_col,
    const char* text,
    PB_FORMAT format)
{
    if (!worksheet || !text) return -1;

    std::string utf8 = ansi_to_utf8(text);

    return worksheet_merge_range(
        safe_ws(worksheet),
        first_row,
        first_col,
        last_row,
        last_col,
        utf8.c_str(),
        safe_fmt(format)
    );
}


// ------------------------------------------------------------
//  Insert image
// ------------------------------------------------------------
__declspec(dllexport) int __stdcall pb_worksheet_insert_image(
    PB_WORKSHEET worksheet,
    std::int32_t row,
    std::int32_t col,
    const char* filename)
{
    if (!worksheet || !filename) return -1;

    std::string utf8 = ansi_to_utf8(filename);

    return worksheet_insert_image(
        safe_ws(worksheet),
        row,
        col,
        utf8.c_str()
    );
}


// ------------------------------------------------------------
//  Autofilter
// ------------------------------------------------------------
__declspec(dllexport) int __stdcall pb_worksheet_autofilter(
    PB_WORKSHEET worksheet,
    std::int32_t first_row,
    std::int32_t first_col,
    std::int32_t last_row,
    std::int32_t last_col)
{
    if (!worksheet) return -1;
    return worksheet_autofilter(safe_ws(worksheet), first_row, first_col, last_row, last_col);
}


// ------------------------------------------------------------
//  Close workbook
// ------------------------------------------------------------
__declspec(dllexport) int __stdcall pb_workbook_close(PB_WORKBOOK workbook) {
    if (!workbook) return -1;
    return workbook_close((lxw_workbook*)workbook);
}


// ------------------------------------------------------------
//  Autofit helper
// ------------------------------------------------------------
__declspec(dllexport)
int __stdcall pb_worksheet_autofit_column(
    PB_WORKSHEET ws,
    std::int32_t col,
    int max_chars,
    PB_FORMAT format)
{
    lxw_worksheet* w = safe_ws(ws);
    if (!w) return -1;

    double width = (max_chars > 0 ? max_chars + 1.0 : 8.43);

    return worksheet_set_column(w, col, col, width, safe_fmt(format));
}

extern "C" __declspec(dllexport)
void __stdcall pb_worksheet_freeze_panes(lxw_worksheet* ws, uint32_t row, uint32_t col)
{
    worksheet_freeze_panes(ws, row, col);
}

} // extern "C"
