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

#ifndef __XLS_WORKSHEET_H__
#define __XLS_WORKSHEET_H__

#include <stdio.h>

#include "biffwriter.h"
#include "bsdqueue.h"
#include "format.h"

struct col_info {
  int first_col;
  int last_col;
  int col_width;
  int xf;
  int grbit;

  TAILQ_ENTRY(col_info) cis;
};

struct wsheetctx {
  struct bwctx base;
  char *name;
  int index;
  int activesheet;
  int firstsheet;
  struct xl_format *url_format;
  int using_tmpfile;

  FILE *fp;
  int fileclosed;
  int offset;
  int xls_rowmax;
  int xls_colmax;
  int xls_strmax;
  int dim_rowmin;
  int dim_rowmax;
  int dim_colmin;
  int dim_colmax;

  int sel_frow;
  int sel_fcol;
  int sel_lrow;
  int sel_lcol;

  TAILQ_HEAD(colinfo_list, col_info) colinfos;
};

struct wsheetctx * wsheet_new(char *name, int index, int activesheet, int firstsheet, struct xl_format *url, int store_in_memory);
void wsheet_destroy(struct wsheetctx *xls);
int xls_write_number(struct wsheetctx *xls, int row, int col, double num);
int xls_write_string(struct wsheetctx *xls, int row, int col, char *str);
int xls_writef_string(struct wsheetctx *xls, int row, int col, char *str, struct xl_format *fmt);
int xls_writef_number(struct wsheetctx *xls, int row, int col, double num, struct xl_format *fmt);
int xls_write_blank(struct wsheetctx *xls, int row, int col, struct xl_format *fmt);
int wsheet_write_url(struct wsheetctx *wsheet, int row, int col, char *url, char *str, struct xl_format *fmt);
void wsheet_close(struct wsheetctx *xls);
unsigned char *wsheet_get_data(struct wsheetctx *ws, size_t *sz);
void wsheet_set_column(struct wsheetctx *ws, int fcol, int lcol, int width);
void wsheet_set_selection(struct wsheetctx *ws, int frow, int fcol, int lrow, int lcol);
void wsheet_set_row(struct wsheetctx *ws, int row, int height, struct xl_format *fmt);

#endif /* __XLS_WORKSHEET_H__ */
