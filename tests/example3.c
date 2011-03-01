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

int main(int argc, char *argv[])
{
  struct wbookctx *wbook;
  struct wsheetctx *one;
  struct xl_format *fmt;

  /* Create a new Excel Workbook */
  wbook = wbook_new("example3.xls", 0);
  one = wbook_addworksheet(wbook, "Formula");

  fmt = wbook_addformat(wbook);
  fmt_set_bold(fmt, 1);
  fmt_set_color(fmt, "blue");

  xls_writef_string(one, 0, 0, "Formula examples", fmt);
  wsheet_writef_formula(one, 1, 1, "=2+3*4", NULL);

  wbook_close(wbook);
  wbook_destroy(wbook);

  return 0;
}
