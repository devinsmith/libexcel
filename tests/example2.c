/*
 * Copyright (c) 2010 Devin Smith <devin@devinsmith.net>
 *
 * Permission to use, copy, modify, and distribute this software for any
 * purpose with or without fee is hereby granted, provided that the above
 * copyright notice and this permission notice appear in all copies.
 *
 * THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
 * WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
 * MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
 * ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
 * WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
 * ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
 * OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "excel.h"

struct wbookctx *gwbook;
struct xl_format *gheading, *gcenter;

void intro(void)
{
  struct wsheetctx *one;
  struct xl_format *fmt;

  /* Create a new Excel Workbook */
  one = wbook_addworksheet(gwbook, "Introduction");
  wsheet_set_column(one, 0, 0, 60);

  fmt = wbook_addformat(gwbook);
  fmt_set_bold(fmt, 1);
  fmt_set_color(fmt, "blue");
  fmt_set_align(fmt, "center");

  xls_writef_string(one, 2, 0, "This workbook demonstrates some of", fmt);
  xls_writef_string(one, 3, 0, "the formatting options provided by", fmt);
  xls_writef_string(one, 4, 0, "the libexcel library.", fmt);
}

struct fsfv {
  int size;
  char *name;
};

void fonts(void)
{
  struct wsheetctx *sheet;
  struct fsfv fonts[] = {
    {10, "Arial"},
    {12, "Arial"},
    {14, "Arial"},
    {12, "Arial Black"},
    {12, "Arial Narrow"},
    {12, "Century Schoolbook"},
    {12, "Courier"},
    {12, "Courier New"},
    {12, "Garamond"},
    {12, "Impact"},
    {12, "Lucida Handwriting"},
    {12, "Times New Roman"},
    {12, "Symbol"},
    {12, "Wingdings"},
    {12, "Calibri"},
    {12, "A font that does not exist"}
  };
  int i;
  int nfs;

  sheet = wbook_addworksheet(gwbook, "Fonts");
  wsheet_set_column(sheet, 0, 0, 30);
  wsheet_set_column(sheet, 1, 1, 10);

  xls_writef_string(sheet, 0, 0, "Font name", gheading);
  xls_writef_string(sheet, 0, 1, "Font size", gheading);

  nfs = sizeof(fonts) / (sizeof(struct fsfv));
  for (i = 0; i < nfs; i++) {
    struct xl_format *fmt;

    fmt = wbook_addformat(gwbook);
    fmt_set_size(fmt, fonts[i].size);
    fmt_set_font(fmt, fonts[i].name);
    xls_writef_string(sheet, i+1, 0, fonts[i].name, fmt);
    xls_writef_number(sheet, i+1, 1, fonts[i].size, fmt);
  }
}

void named_colors(void)
{
  struct wsheetctx *sheet;
  int i;

  char *colors[] = { "aqua", "black", "blue", "fuchsia", "gray", "green",
    "lime", "navy", "orange", "purple", "red", "silver", "white", "yellow" };
  int indices[] = { 0x0F, 0x08, 0x0C, 0x0E, 0x17, 0x11, 0x0B, 0x12, 0x1D,
    0x24, 0x0A, 0x16, 0x09, 0x0D };

  sheet = wbook_addworksheet(gwbook, "Named colors");
  wsheet_set_column(sheet, 0, 3, 15);
  xls_writef_string(sheet, 0, 0, "Index", gheading);
  xls_writef_string(sheet, 0, 1, "Index", gheading);
  xls_writef_string(sheet, 0, 2, "Name", gheading);
  xls_writef_string(sheet, 0, 3, "Color", gheading);

  for (i = 0; i < 14; i++) {
    struct xl_format *fmt;
    char fmt_name[12];

    fmt = wbook_addformat(gwbook);
    fmt_set_color(fmt, colors[i]);
    fmt_set_align(fmt, "center");

    snprintf(fmt_name, sizeof(fmt_name), "0x%02X", indices[i]);

    xls_writef_number(sheet, i+1, 0, indices[i], gcenter);
    xls_writef_string(sheet, i+1, 1, fmt_name, gcenter);
    xls_writef_string(sheet, i+1, 2, colors[i], gcenter);
    xls_writef_string(sheet, i+1, 3, colors[i], fmt);
  }
}

void standard_colors(void)
{
  struct wsheetctx *sheet;
  int i;

  sheet = wbook_addworksheet(gwbook, "Standard colors");
  wsheet_set_column(sheet, 0, 3, 15);
  xls_writef_string(sheet, 0, 0, "Index", gheading);
  xls_writef_string(sheet, 0, 1, "Index", gheading);
  xls_writef_string(sheet, 0, 2, "Color", gheading);

  for (i = 8; i < 64; i++) {
    struct xl_format *fmt;
    char fmt_name[12];

    fmt = wbook_addformat(gwbook);
    fmt_set_colori(fmt, i);
    fmt_set_align(fmt, "center");

    snprintf(fmt_name, sizeof(fmt_name), "0x%02X", i);
    xls_writef_number(sheet, i - 7, 0, i, gcenter);
    xls_writef_string(sheet, i - 7, 1, fmt_name, gcenter);
    xls_writef_string(sheet, i - 7, 2, "COLOR", fmt);
  }
}

struct numfmt {
  int index;
  double val;
  double nval;
  char *desc;
};

void numeric_formats(void)
{
  struct wsheetctx *sheet;
  int i, nfs;
  struct numfmt numfmts[] = {
    { 0x00, 1234.567, 0, "General" },
    { 0x01, 1234.567, 0, "0" },
    { 0x02, 1234.567, 0, "0.00" },
    { 0x03, 1234.567, 0, "#,##0" },
    { 0x04, 1234.567, 0, "#,##0.00" },
    { 0x05, 1234.567, -1234.567, "($#,##0_);($#,##0)" },
    { 0x06, 1234.567, -1234.567, "($#,##0_);[Red]($#,##0)" },
    { 0x07, 1234.567, -1234.567, "($#,##0.00_);($#,##0.00)" },
    { 0x08, 1234.567, -1234.567, "($#,##0.00_);[Red]($#,##0.00)" },
    { 0x09, 0.567, 0, "0%" },
    { 0x0A, 0.567, 0, "0.00%" },
    { 0x0B, 1234.567, 0, "0.00E+00" },
    { 0x0C, 0.75, 0, "# ?/?" },
    { 0x0D, 0.3125, 0, "# \?\?/??" },
    { 0x0E, 36870.016, 0, "m/d/yy" },
    { 0x0F, 36870.016, 0, "d-mmm-yy" },
    { 0x10, 36870.016, 0, "d-mmm" },
    { 0x11, 36870.016, 0, "mmm-yy" },
    { 0x12, 36870.016, 0, "h:mm AM/PM" },
    { 0x13, 36870.016, 0, "h:mm:ss AM/PM" },
    { 0x14, 36870.016, 0, "h:mm" },
    { 0x15, 36870.016, 0, "h:mm:ss" },
    { 0x16, 36870.016, 0, "m/d/yy h:mm" },
    { 0x25, 1234.567, -1234.567, "(#,##0_);(#,##0)" },
    { 0x26, 1234.567, -1234.567, "(#,##0_);[Red](#,##0)" },
    { 0x27, 1234.567, -1234.567, "(#,##0.00_);(#,##0.00)" },
    { 0x28, 1234.567, -1234.567, "(#,##0.00_);[Red](#,##0.00)" },
    { 0x29, 1234.567, -1234.567, "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)" },
    { 0x2A, 1234.567, -1234.567, "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)" },
    { 0x2B, 1234.567, -1234.567, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"\?\?_);_(@_)" },
    { 0x2C, 1234.567, -1234.567, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"\?\?_);_(@_)" },
    { 0x2D, 36870.016, 0, "mm:ss" },
    { 0x2E, 3.0153, 0, "[h]:mm:ss" },
    { 0x2F, 36870.016, 0, "mm:ss.0" },
    { 0x30, 1234.567, 0, "##0.0E+0" },
    { 0x31, 1234.567, 0, "@" }
  };

  sheet = wbook_addworksheet(gwbook, "Numeric formats");
  wsheet_set_column(sheet, 0, 4, 15);
  wsheet_set_column(sheet, 5, 5, 45);

  xls_writef_string(sheet, 0, 0, "Index", gheading);
  xls_writef_string(sheet, 0, 1, "Index", gheading);
  xls_writef_string(sheet, 0, 2, "Unformatted", gheading);
  xls_writef_string(sheet, 0, 3, "Formatted", gheading);
  xls_writef_string(sheet, 0, 4, "Negative", gheading);
  xls_writef_string(sheet, 0, 5, "Format", gheading);

  nfs = sizeof(numfmts) / sizeof(struct numfmt);
  for (i = 0; i < nfs; i++) {
    struct xl_format *fmt;
    char fmt_name[12];

    fmt = wbook_addformat(gwbook);
    fmt_set_num_format(fmt, numfmts[i].index);

    snprintf(fmt_name, sizeof(fmt_name), "0x%02X", numfmts[i].index);
    xls_writef_number(sheet, i + 1, 0, numfmts[i].index, gcenter);
    xls_writef_string(sheet, i + 1, 1, fmt_name, gcenter);
    xls_writef_number(sheet, i + 1, 2, numfmts[i].val, gcenter);
    xls_writef_number(sheet, i + 1, 3, numfmts[i].val, fmt);

    if (numfmts[i].nval)
      xls_writef_number(sheet, i + 1, 4, numfmts[i].nval, fmt);

    xls_write_string(sheet, i + 1, 5, numfmts[i].desc);
  }
}

void borders(void)
{
  struct wsheetctx *sheet;
  int i;

  sheet = wbook_addworksheet(gwbook, "Borders");

  wsheet_set_column(sheet, 0, 4, 10);
  wsheet_set_column(sheet, 5, 5, 40);

  xls_writef_string(sheet, 0, 0, "Index", gheading);
  xls_writef_string(sheet, 0, 1, "Index", gheading);
  xls_writef_string(sheet, 0, 3, "Style", gheading);
  xls_writef_string(sheet, 0, 5, "The style is highlighted in red for ", gheading);
  xls_writef_string(sheet, 1, 5, "emphasis, the default color is black.", gheading);

  for (i = 0; i < 8; i++) {
    struct xl_format *fmt;
    char fmt_name[12];

    fmt = wbook_addformat(gwbook);
    fmt_set_border(fmt, i);
    fmt_set_border_color(fmt, "red");
    fmt_set_align(fmt, "center");

    snprintf(fmt_name, sizeof(fmt_name), "0x%02X", i);
    xls_writef_number(sheet, 2 * (i + 1), 0, i, gcenter);
    xls_writef_string(sheet, 2 * (i + 1), 1, fmt_name, gcenter);
    xls_writef_string(sheet, 2 * (i + 1), 3, "Border", fmt);
  }
}

void patterns(void)
{
  struct wsheetctx *sheet;
  int i;

  sheet = wbook_addworksheet(gwbook, "Patterns");

  wsheet_set_column(sheet, 0, 4, 10);
  wsheet_set_column(sheet, 5, 5, 50);

  xls_writef_string(sheet, 0, 0, "Index", gheading);
  xls_writef_string(sheet, 0, 1, "Index", gheading);
  xls_writef_string(sheet, 0, 3, "Pattern", gheading);
  xls_writef_string(sheet, 0, 5, "The background color has been set to silver.", gheading);
  xls_writef_string(sheet, 1, 5, "The foreground color has been set to green.", gheading);

  for (i = 0; i < 32; i++) {
    struct xl_format *fmt;
    char fmt_name[12];

    fmt = wbook_addformat(gwbook);
    fmt_set_pattern(fmt, i);
    fmt_set_bg_color(fmt, "silver");
    fmt_set_fg_color(fmt, "green");
    fmt_set_align(fmt, "center");

    snprintf(fmt_name, sizeof(fmt_name), "0x%02X", i);
    xls_writef_number(sheet, 2 * (i + 1), 0, i, gcenter);
    xls_writef_string(sheet, 2 * (i + 1), 1, fmt_name, gcenter);
    xls_writef_string(sheet, 2 * (i + 1), 3, "Pattern", fmt);

    if (i == 1) {
      xls_writef_string(sheet, 2 * (i + 1), 5, "This is solid color, the most useful pattern", gheading);

    }
  }
}

void alignment(void)
{
  struct wsheetctx *sheet;
  struct xl_format *fmt[15];
  int i;

  sheet = wbook_addworksheet(gwbook, "Alignment");

  wsheet_set_column(sheet, 0, 7, 12);
  //wsheet_set_row(sheet, 0, 40, NULL);
  wsheet_set_selection(sheet, 7, 0, 0, 0);

  for (i = 0; i < 15; i++) {
    fmt[i] = wbook_addformat(gwbook);
  }

  fmt_set_align(fmt[1], "top");
  fmt_set_align(fmt[2], "bottom");
  fmt_set_align(fmt[3], "vcenter");
  fmt_set_align(fmt[4], "vjustify");
  fmt_set_text_wrap(fmt[5], 1);

  fmt_set_align(fmt[6], "left");
  fmt_set_align(fmt[7], "right");
  fmt_set_align(fmt[8], "center");
  fmt_set_align(fmt[9], "fill");
  fmt_set_align(fmt[10], "justify");
  fmt_set_merge(fmt[11]);

  fmt_set_rotation(fmt[12], 1);
  fmt_set_rotation(fmt[13], 2);
  fmt_set_rotation(fmt[14], 3);

  xls_writef_string(sheet, 0, 0, "Vertical", gheading);
  xls_writef_string(sheet, 0, 1, "top", fmt[1]);
  xls_writef_string(sheet, 0, 2, "bottom", fmt[2]);
  xls_writef_string(sheet, 0, 3, "vcenter", fmt[3]);
  xls_writef_string(sheet, 0, 4, "vjustify", fmt[4]);
  xls_writef_string(sheet, 0, 5, "text\nwrap", fmt[5]);

  xls_writef_string(sheet, 2, 0, "Horizontal", gheading);
  xls_writef_string(sheet, 2, 1, "left", fmt[6]);
  xls_writef_string(sheet, 2, 2, "right", fmt[7]);
  xls_writef_string(sheet, 2, 3, "center", fmt[8]);
  xls_writef_string(sheet, 2, 4, "fill", fmt[9]);
  xls_writef_string(sheet, 2, 5, "justify", fmt[10]);

  xls_writef_string(sheet, 2, 6, "merge", fmt[11]);
  xls_write_blank(sheet, 2, 7, fmt[11]);

  xls_writef_string(sheet, 4, 0, "Rotation", gheading);
  xls_writef_string(sheet, 4, 1, "Rotate 1", fmt[12]);
  xls_writef_string(sheet, 4, 2, "Rotate 2", fmt[13]);
  xls_writef_string(sheet, 4, 3, "Rotate 3", fmt[14]);
}

int main(int argc, char *argv[])
{
  /* Create a new Excel Workbook */
  gwbook = wbook_new("example2.xls", 0);

  /* Some common formats */
  gheading = wbook_addformat(gwbook);
  gcenter = wbook_addformat(gwbook);
  fmt_set_align(gcenter, "center");
  fmt_set_align(gheading, "center");
  fmt_set_bold(gheading, 1);

  intro();
  fonts();
  named_colors();
  standard_colors();
  numeric_formats();
  borders();
  patterns();
  alignment();

  /* This is required */
  wbook_close(gwbook);
  wbook_destroy(gwbook);

  return 0;
}
