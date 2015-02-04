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

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "workbook.h"
#include "stream.h"
#include "hashhelp.h"

void wbook_store_window1(struct wbookctx *wbook);
void wbook_store_all_fonts(struct wbookctx *wbook);
void wbook_store_all_xfs(struct wbookctx *wbook);
void wbook_store_all_styles(struct wbookctx *wbook);
void wbook_store_boundsheet(struct wbookctx *wbook, char *sname, int offset);
void wbook_store_workbook(struct wbookctx *wbook);
static void wbook_store_1904(struct wbookctx *wbook);
static void wbook_store_num_format(struct wbookctx *wbook, char *format, int index);
static void wbook_store_codepage(struct wbookctx *wbook);
void wbook_store_all_num_formats(struct wbookctx *wbook);

struct wbookctx *wbook_new(char *filename, int store_in_memory)
{
	return wbook_new_ex(xl_file_handler, filename, store_in_memory);
}

struct wbookctx *wbook_new_ex(struct xl_io_handler io_handler, char *filename, int store_in_memory)
{
  struct wbookctx *wbook;

  wbook = malloc(sizeof(struct wbookctx));
  wbook->biff = bw_new();
  wbook->OLEwriter = ow_new_ex(io_handler,filename);
  if (wbook->OLEwriter == NULL) {
    free(wbook);
    return NULL;
  }
  wbook->store_in_memory = store_in_memory;
  wbook->epoch1904 = 0;
  wbook->activesheet = 0;
  wbook->firstsheet = 0;
  wbook->xf_index = 16;  /* 15 style XF's and 1 cell XF. */
  wbook->fileclosed = 0;
  wbook->biffsize = 0;
  wbook->sheetname = "Sheet";
  wbook->tmp_format = fmt_new(0);
  wbook->url_format = NULL;
  wbook->codepage = 0x04E4; /* 1252 */
  wbook->sheets = NULL;
  wbook->sheetcount = 0;
  wbook->formats = NULL;
  wbook->formatcount = 0;

  /* Add the default format for hyperlinks */
  wbook->url_format = wbook_addformat(wbook);
  fmt_set_fg_color(wbook->url_format, "blue");
  fmt_set_underline(wbook->url_format, 1);

  return wbook;
}

void wbook_close(struct wbookctx *wbook)
{
  if (wbook->fileclosed)
    return;

  wbook_store_workbook(wbook);
  ow_close(wbook->OLEwriter);
  wbook->fileclosed = 1;
}

void wbook_destroy(struct wbookctx *wbook)
{
  int i;

  if (!wbook->fileclosed)
    wbook_close(wbook);

  /* Delete all sheets */
  for (i = 0; i < wbook->sheetcount; i++) {
    wsheet_destroy(wbook->sheets[i]);
  }

  /* Delete all formats */
  for (i = 0; i < wbook->formatcount; i++) {
    free(wbook->formats[i]);
  }

  fmt_destroy(wbook->tmp_format);
  free(wbook->sheets);
  free(wbook);
}

struct wsheetctx * wbook_addworksheet(struct wbookctx *wbook, char *sname)
{
  char *name = sname;
  int index;
  struct wsheetctx *wsheet;

  if (name && strlen(name) > 31)
    name[31] = '\0';

  index = wbook->sheetcount;
  if (sname == NULL)
  {
    int len = strlen(wbook->sheetname) + 20;

    name = malloc(len);
    snprintf(name, len, "%s%d", wbook->sheetname, index + 1);
  }

  if (wbook->sheets == NULL)
    wbook->sheets = malloc(sizeof(struct wsheetctx *));
  else
    wbook->sheets = realloc(wbook->sheets, sizeof(struct wsheetctx *) * (index + 1));

  wsheet = wsheet_new(name, index, wbook->activesheet, wbook->firstsheet,
      wbook->url_format, wbook->store_in_memory);
  wbook->sheets[index] = wsheet;
  wbook->sheetcount++;

  return wsheet;
}

/* Add a new format to the Excel workbook.  This adds an XF record and
 * a FONT record.  TODO: add a FORMAT record. */
struct xl_format *wbook_addformat(struct wbookctx *wbook)
{
  int index;
  struct xl_format *fmt;

  index = wbook->formatcount;

  if (wbook->formats == NULL)
    wbook->formats = malloc(sizeof(struct xl_format *));
  else
    wbook->formats = realloc(wbook->formats, sizeof(struct xl_format *) * (index + 1));

  fmt = fmt_new(wbook->xf_index);
  wbook->xf_index += 1;

  wbook->formats[index] = fmt;
  wbook->formatcount++;

  return fmt;
}

/****************************************************************************
 *
 * _calc_sheet_offsets()
 *
 * Calculate offsets for Worksheet BOF records. */
void wbook_calc_sheet_offsets(struct wbookctx *wbook)
{
  int oBOF = 11;
  int oEOF = 4;
  int offset = wbook->biff->datasize;
  int i;

  for (i = 0; i < wbook->sheetcount; i++) {
    offset += oBOF + strlen(wbook->sheets[i]->name);
  }
  offset += oEOF;

  for (i = 0; i < wbook->sheetcount; i++) {
      wbook->sheets[i]->offset = offset;
      offset += ((struct bwctx *)wbook->sheets[i])->datasize;
  }

  wbook->biffsize = offset;
}

/*
 * wbook_store_workbook(struct wbookctx *wbook)
 *
 * Assemble worksheets into a workbook and send the BIFF data to an OLE
 * storage.
 *
 */
void wbook_store_workbook(struct wbookctx *wbook)
{
  struct owctx *ole = wbook->OLEwriter;
  int i;

  /* Call the finalization methods for each worksheet */
  for (i = 0; i < wbook->sheetcount; i++) {
    wsheet_close(wbook->sheets[i]);
  }

  /* Add workbook globals */
  bw_store_bof(wbook->biff, 0x0005);
  wbook_store_codepage(wbook);
  wbook_store_window1(wbook);
  wbook_store_1904(wbook);
  wbook_store_all_fonts(wbook);
  wbook_store_all_num_formats(wbook);
  wbook_store_all_xfs(wbook);
  wbook_store_all_styles(wbook);
  wbook_calc_sheet_offsets(wbook);

  /* Add BOUNDSHEET records */
  for (i = 0; i < wbook->sheetcount; i++) {
    wbook_store_boundsheet(wbook, wbook->sheets[i]->name, wbook->sheets[i]->offset);
  }

  bw_store_eof(wbook->biff);

  /* Write Worksheet data if data <~ 7MB */
  if (ow_set_size(ole, wbook->biffsize)) {
    ow_write_header(ole);
    ow_write(ole, wbook->biff->data, wbook->biff->datasize);

    for (i = 0; i < wbook->sheetcount; i++) {
      unsigned char *tmp;
      size_t size;

      while ((tmp = wsheet_get_data(wbook->sheets[i], &size))) {
        ow_write(ole, tmp, size);
        free(tmp);
      }
    }
  }
}

/* Write Excel BIFF5-8 WINDOW1 record. */
void wbook_store_window1(struct wbookctx *wbook)
{
  struct pkt *pkt;

  pkt = pkt_init(20, VARIABLE_PACKET);

  /* Write header */
  pkt_add16_le(pkt, 0x003D);  /* Record identifier */
  pkt_add16_le(pkt, 0x0012);  /* Number of bytes to follow */

  /* Write data */
  pkt_add16_le(pkt, 0x0000);  /* Horizontal position of window */
  pkt_add16_le(pkt, 0x0069);  /* Vertical position of window */
  pkt_add16_le(pkt, 0x339F);  /* Width of window */
  pkt_add16_le(pkt, 0x5D1B);  /* Height of window */
  pkt_add16_le(pkt, 0x0038);  /* Option flags */
  pkt_add16_le(pkt, wbook->activesheet);  /* Selected worksheet */
  pkt_add16_le(pkt, wbook->firstsheet);   /* 1st displayed worksheet */
  pkt_add16_le(pkt, 0x0001);  /* Number of workbook tabs selected */
  pkt_add16_le(pkt, 0x0258);  /* Tab to scrollbar ratio */

  bw_append(wbook->biff, pkt->data, pkt->len);
  pkt_free(pkt);
}

/* Write all FONT records. */
void wbook_store_all_fonts(struct wbookctx *wbook)
{
  int i;
  struct pkt *font;
  struct htbl *fonts;
  int key;
  int index;

  font = fmt_get_font(wbook->tmp_format);
  for (i = 1; i < 6; i++) {
    bw_append(wbook->biff, font->data, font->len);
  }
  pkt_free(font);

  fonts = hashtbl_new(wbook->formatcount + 1); /* For tmp_format */
  index = 6;  /* First user defined FONT */

  key = fmt_gethash(wbook->tmp_format);
  hashtbl_insert(fonts, key, 0);  /* Index of the default font */

  /* User defined fonts */
  for (i = 0; i < wbook->formatcount; i++) {
    int data;
    key = fmt_gethash(wbook->formats[i]);
    data = hashtbl_get(fonts, key);
    if (data >= 0) {
      /* FONT has already been used */
      wbook->formats[i]->font_index = data;
    } else {
      /* Add a new FONT record */
      hashtbl_insert(fonts, key, index);
      wbook->formats[i]->font_index = index;
      index++;
      font = fmt_get_font(wbook->formats[i]);
      bw_append(wbook->biff, font->data, font->len);
      pkt_free(font);
    }
  }

  hashtbl_destroy(fonts);
}

/* Store user defined numerical formats ie. FORMAT records */
void wbook_store_all_num_formats(struct wbookctx *wbook)
{
  int index = 164;  /* Start from 0xA4 */
  struct htbl *num_formats;
  int key;
  int i;

  num_formats = hashtbl_new(1);

  /* User defined formats */
  for (i = 0; i < wbook->formatcount; i++) {
    int data;
    if (wbook->formats[i]->num_format_str == NULL)
      continue;
    key = fmt_gethash(wbook->formats[i]);
    data = hashtbl_get(num_formats, key);
    if (data >= 0) {
      /* FONT has already been used */
      wbook->formats[i]->num_format = data;
    } else {
      /* Add a new FONT record */
      hashtbl_insert(num_formats, key, index);
      wbook->formats[i]->num_format = index;
      wbook_store_num_format(wbook, wbook->formats[i]->num_format_str, index);
      index++;
    }
  }
  hashtbl_destroy(num_formats);
}

/* Write all XF records */
void wbook_store_all_xfs(struct wbookctx *wbook)
{
  int i;
  struct pkt *xf;

  xf = fmt_get_xf(wbook->tmp_format, 0xFFF5); /* Style XF */
  for (i = 0; i <= 14; i++) {
    bw_append(wbook->biff, xf->data, xf->len);
  }
  pkt_free(xf);

  xf = fmt_get_xf(wbook->tmp_format, 0x0001); /* Cell XF */
  bw_append(wbook->biff, xf->data, xf->len);
  pkt_free(xf);

  /* User defined formats */
  for (i = 0; i < wbook->formatcount; i++) {
    xf = fmt_get_xf(wbook->formats[i], 0x0001);
    bw_append(wbook->biff, xf->data, xf->len);
    pkt_free(xf);
  }
}

/* Write Excel BIFF STYLE record */
void wbook_store_style(struct wbookctx *wbook)
{
  uint16_t name = 0x0093;  /* Record identifier */
  uint16_t length = 0x0004; /* Bytes to follow */
  uint16_t ixfe = 0x0000;

  struct pkt *pkt;

  pkt = pkt_init(8, FIXED_PACKET);

  /* Write header */
  pkt_add16_le(pkt, name);
  pkt_add16_le(pkt, length);

  /* Write data */
  pkt_add16_le(pkt, ixfe);  /* Index to style XF */
  pkt_add8(pkt, 0x00); /* Built-in style */
  pkt_add8(pkt, 0x00); /* Outline style level */

  bw_append(wbook->biff, pkt->data, pkt->len);
  pkt_free(pkt);

}

/* Write all STYLE records. */
void wbook_store_all_styles(struct wbookctx *wbook)
{
  wbook_store_style(wbook);
}

/* Writes Excel BIFF BOUNDSHEET record */
void wbook_store_boundsheet(struct wbookctx *wbook, char *sname, int offset)
{
  uint16_t name = 0x0085;  /* Record identifier */
  uint16_t length;
  uint16_t grbit = 0x0000;
  int cch;
  struct pkt *pkt;

  length = 0x07 + strlen(sname);  /* Number of bytes to follow */
  cch = strlen(sname);

  pkt = pkt_init(0, VARIABLE_PACKET);

  /* Write header */
  pkt_add16_le(pkt, name);
  pkt_add16_le(pkt, length);

  /* Write data */
  pkt_add32_le(pkt, offset);  /* Location of worksheet BOF */
  pkt_add16_le(pkt, grbit);   /* Sheet identifier */
  pkt_add8(pkt, cch); /* Length of sheet name */
  pkt_addraw(pkt, (unsigned char *)sname, cch);

  bw_append(wbook->biff, pkt->data, pkt->len);
  pkt_free(pkt);

}

/* Writes Excel FORMAT record for non "built-in" numerical formats. */
static void wbook_store_num_format(struct wbookctx *wbook, char *format, int index)
{
  struct pkt *pkt;
  size_t cch;

  cch = strlen(format);
  pkt = pkt_init(0, VARIABLE_PACKET);

  /* Write header */
  pkt_add16_le(pkt, 0x041E);  /* Record identifier */
  pkt_add16_le(pkt, 0x0003 + cch);

  /* Write data */
  pkt_add16_le(pkt, index);
  pkt_add8(pkt, cch);
  pkt_addraw(pkt, (unsigned char *)format, cch);

  bw_append(wbook->biff, pkt->data, pkt->len);

  pkt_free(pkt);
}

/* Write Excel 1904 record to indicate the date system in use. */
static void wbook_store_1904(struct wbookctx *wbook)
{
  struct pkt *pkt;

  pkt = pkt_init(6, FIXED_PACKET);

  /* Write header */
  pkt_add16_le(pkt, 0x0022);  /* Record identifier */
  pkt_add16_le(pkt, 0x0002);

  /* Write data */
  pkt_add16_le(pkt, wbook->epoch1904);  /* Flag for 1904 date system */

  bw_append(wbook->biff, pkt->data, pkt->len);
  pkt_free(pkt);
}

/* Stores the CODEPAGE biff record */
static void wbook_store_codepage(struct wbookctx *wbook)
{
  struct pkt *pkt;

  pkt = pkt_init(6, FIXED_PACKET);

  /* Write header */
  pkt_add16_le(pkt, 0x0042);  /* Record identifier */
  pkt_add16_le(pkt, 0x0002);

  /* Write data */
  pkt_add16_le(pkt, wbook->codepage);

  bw_append(wbook->biff, pkt->data, pkt->len);
  pkt_free(pkt);
}
