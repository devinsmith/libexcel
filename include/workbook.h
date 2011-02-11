/*
 * Copyright (c) 2010 Devin Smith <devin@devinsmith.net>
 * Based on early versions of Spreadsheet::WriteExcel
 * Copyright (c) 2000-2001 John McNamara <jmcnamara@cpan.org>
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

#ifndef __XLS_WORKBOOK_H__
#define __XLS_WORKBOOK_H__

#include <stdio.h>

#include "biffwriter.h"
#include "format.h"
#include "olewriter.h"
#include "worksheet.h"
#include "bsdqueue.h"

struct wbookctx {
  struct bwctx *biff;

  int store_in_memory;
  struct owctx *OLEwriter;
  int epoch1904;
  int activesheet;
  int firstsheet;
  int fileclosed;
  int biffsize;
  int codepage;
  char *sheetname;
  struct xl_format *tmp_format;
  struct xl_format *url_format;
  int xf_index;

  int sheetcount;
  struct wsheetctx **sheets;

  int formatcount;
  struct xl_format **formats;
};

struct wbookctx *wbook_new(char *filename, int store_in_memory);
void wbook_close(struct wbookctx *wb);
void wbook_destroy(struct wbookctx *wb);
struct wsheetctx *wbook_addworksheet(struct wbookctx *wbook, char *sname);
struct xl_format *wbook_addformat(struct wbookctx *wbook);

#endif /* __XLS_WORKBOOK_H__ */
