/*
 * Copyright (c) 2010 Devin Smith <devin@devinsmith.net>
 * Based on early versions of Spreadsheet::WriteExcel
 * Copyright (c) 2000 John McNamara <jmcnamara@cpan.org>
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

#include "worksheet.h"
#include "stream.h"

#define XLS_ROWMAX 65536
#define XLS_COLMAX 256
#define XLS_STRMAX 255

int xls_init(struct wsheetctx *xls, char *name, int index, int activesheet, int firstsheet, struct xl_format *url, int store_in_memory);
void wsheet_store_dimensions(struct wsheetctx *xls);
void wsheet_store_window2(struct wsheetctx *xls);
void wsheet_store_selection(struct wsheetctx *xls, int frow, int fcol, int lrow, int lcol);
void wsheet_store_colinfo(struct wsheetctx *wsheet, struct col_info *ci);
void wsheet_store_defcol(struct wsheetctx *wsheet);
void wsheet_append(void *xlsctx, void *data, size_t sz);

extern int bw_init(struct bwctx *bw);

#ifdef WIN32
#include <io.h>
#include <windows.h>
#define tmpfile() _xls_win32_tmpfile()

FILE *_xls_win32_tmpfile(void)
{
  char path_name[MAX_PATH + 1];
  char file_name[MAX_PATH + 1];
  DWORD path_len;
  HANDLE handle;
  int fd;
  FILE *fp;


  path_len = GetTempPath(MAX_PATH, path_name);
  if (path_len <= 0 || path_len >= MAX_PATH)
    return NULL;

  if (GetTempFileName(path_name, "xl_", 0, file_name) == 0)
    return NULL;

  handle = CreateFile(file_name, GENERIC_READ | GENERIC_WRITE,
      0, NULL, CREATE_ALWAYS,
      FILE_ATTRIBUTE_NORMAL | FILE_FLAG_DELETE_ON_CLOSE,
      NULL);

  if (handle == INVALID_HANDLE_VALUE) {
    DeleteFile(file_name);
    return NULL;
  }

  fd = _open_osfhandle((intptr_t)handle, 0);
  if (fd < 0) {
    CloseHandle(handle);
    return NULL;
  }

  fp = fdopen(fd, "w+b");
  if (fp == NULL) {
    _close(fd);
    return NULL;
  }

  return fp;
}
#endif

struct wsheetctx * wsheet_new(char *name, int index, int activesheet, int firstsheet, struct xl_format *url, int store_in_memory)
{
  struct wsheetctx *xls;

  xls = malloc(sizeof(struct wsheetctx));
  bw_init((struct bwctx *)xls);
  ((struct bwctx *)xls)->append = wsheet_append;
  TAILQ_INIT(&xls->colinfos);

  if (xls_init(xls, name, index, activesheet, firstsheet, url,
        store_in_memory) == -1) {
    free(xls);
    return NULL;
  }

  return xls;
}

void wsheet_destroy(struct wsheetctx *xls)
{
  struct col_info *ci;
  /* Free the entire tail queue of colinfo records. */
  while ((ci = TAILQ_FIRST(&xls->colinfos))) {
    TAILQ_REMOVE(&xls->colinfos, ci, cis);
    free(ci);
  }

  /* Free up anything else that was allocated */
  free(xls->name);
  free(((struct bwctx *)xls)->data);
  free(xls);
}

int xls_init(struct wsheetctx *xls, char *name, int index, int activesheet,
    int firstsheet, struct xl_format *url, int store_in_memory)
{
  int rowmax = 65536;
  int colmax = 256;
  int strmax = 255;

  xls->name = strdup(name);
  xls->index = index;
  xls->activesheet = activesheet;
  xls->firstsheet = firstsheet;
  xls->url_format = url;
  xls->using_tmpfile = 1;

  xls->fp = NULL;
  xls->fileclosed = 0;
  xls->offset = 0;
  xls->xls_rowmax = rowmax;
  xls->xls_colmax = colmax;
  xls->xls_strmax = strmax;
  xls->dim_rowmin = rowmax + 1;
  xls->dim_rowmax = 0;
  xls->dim_colmin = rowmax + 1;
  xls->dim_colmax = 0;
  xls->sel_frow = 0;
  xls->sel_fcol = 0;
  xls->sel_lrow = 0;
  xls->sel_lcol = 0;

  if (xls->using_tmpfile == 1) {
    xls->fp = tmpfile();
    if (xls->fp == NULL)
      xls->using_tmpfile = 0;
  }

  return 0;
}

void wsheet_close(struct wsheetctx *xls)
{
  /* Prepend in reverse order !! */
  wsheet_store_dimensions(xls);

  /* Prepend the COLINFO records if they exist */
  if (!TAILQ_EMPTY(&xls->colinfos)) {
    struct col_info *ci;
    TAILQ_FOREACH(ci, &xls->colinfos, cis) {
      wsheet_store_colinfo(xls, ci);
    }
    wsheet_store_defcol(xls);
  }

  /* Prepend in reverse order!! */
  bw_store_bof((struct bwctx *)xls, 0x0010);

  /* Append */
  wsheet_store_window2(xls);
  wsheet_store_selection(xls, xls->sel_frow, xls->sel_fcol, xls->sel_lrow, xls->sel_lcol);
  bw_store_eof((struct bwctx *)xls);
}

void wsheet_set_selection(struct wsheetctx *xls, int frow, int fcol, int lrow, int lcol)
{
  xls->sel_frow = frow;
  xls->sel_fcol = fcol;
  xls->sel_lrow = lrow;
  xls->sel_lcol = lcol;
}

/****************************************************************************
 * Writes Excel DIMENSIONS to define the area in which there is data.
 */
void wsheet_store_dimensions(struct wsheetctx *xls)
{
  uint16_t name = 0x0000; /* Record identifier */
  uint16_t length = 0x000A; /* Number of bytes to follow */
  uint16_t reserved = 0x0000; /* Reserved by Excel */
  struct pkt *pkt;

  pkt = pkt_init(14, FIXED_PACKET);

  /* Write header */
  pkt_add16_le(pkt, name);
  pkt_add16_le(pkt, length);

  /* Write data */
  pkt_add16_le(pkt, xls->dim_rowmin);
  pkt_add16_le(pkt, xls->dim_rowmax);
  pkt_add16_le(pkt, xls->dim_colmin);
  pkt_add16_le(pkt, xls->dim_colmax);
  pkt_add16_le(pkt, reserved);
  bw_prepend((struct bwctx *)xls, pkt->data, pkt->len);
  pkt_free(pkt);
}

/****************************************************************************
 * void wsheet_store_window2(struct xlsctx *xls)
 *
 * Write BIFF record WinDow2 */
void wsheet_store_window2(struct wsheetctx *xls)
{
  uint16_t name = 0x023E; /* Record identifier */
  uint16_t length = 0x000A; /* Number of bytes to follow */

  uint16_t grbit = 0x00B6; /* Option flags */
  uint16_t rwTop = 0x0000; /* Top row visible in window */
  uint16_t colLeft = 0x0000; /* Leftmost column visible in window */
  uint32_t rgbHdr = 0x00000000; /* Row/column heading and gridline color */
  struct pkt *pkt;

  pkt = pkt_init(14, FIXED_PACKET);

  if (xls->activesheet == xls->index) {
    grbit = 0x06B6;
  }

  /* Write header */
  pkt_add16_le(pkt, name);
  pkt_add16_le(pkt, length);

  /* Write data */
  pkt_add16_le(pkt, grbit);
  pkt_add16_le(pkt, rwTop);
  pkt_add16_le(pkt, colLeft);
  pkt_add32_le(pkt, rgbHdr);
  wsheet_append(xls, pkt->data, pkt->len);
  pkt_free(pkt);
}

/* Retrieves data from memory in one chunk, or from disk in 4096
 * sized chunks. */
unsigned char *wsheet_get_data(struct wsheetctx *ws, size_t *sz)
{
  unsigned char *tmp;
  struct bwctx *biff = (struct bwctx *)ws;

  /* Return data stored in memory */
  if (biff->data) {
    tmp = malloc(biff->_sz);
    memcpy(tmp, biff->data, biff->_sz);
    *sz = biff->_sz;
    free(biff->data);
    biff->data = NULL;
    if (ws->using_tmpfile == 1) {
      fseek(ws->fp, 0, SEEK_SET);
    }
    return tmp;
  }

  if (ws->using_tmpfile == 1) {
    size_t bytes_read;
    tmp = malloc(4096);
    bytes_read = fread(tmp, 1, 4096, ws->fp);
    *sz = bytes_read;
    if (bytes_read <= 0) {
      free(tmp);
      return NULL;
    }
    return tmp;
  }

  return NULL;
}

int wsheet_xf(struct xl_format *fmt)
{
  if (fmt)
    return fmt->xf_index;

  return 0x0F;
}

void wsheet_append(void *xlsctx, void *data, size_t sz)
{
  struct wsheetctx *xls = (struct wsheetctx *)xlsctx;

  if (!xls->using_tmpfile) {
    bw_append((struct bwctx *)xls, data, sz);
  } else {
    fwrite(data, sz, 1, xls->fp);
    /* TODO determine if below line is really needed */
    ((struct bwctx *)xls)->datasize += sz;
  }
}

/* Write a double to the specified row and column (zero indexed).
 * An integer can be written as a double.  Excel will display an integer.
 * This writes the Excel NUMBER record to the worksheet. (BIFF3-BIFF8) */
int xls_writef_number(struct wsheetctx *xls, int row, int col, double num, struct xl_format *fmt)
{
  uint16_t name = 0x0203; /* Record identifier */
  uint16_t length = 0x000E; /* Number of bytes to follow */
  uint16_t xf; /* The cell format */
  unsigned char xl_double[8];
  struct pkt *pkt;
  struct bwctx *biff = (struct bwctx *)xls;

  if (row >= xls->xls_rowmax) { return -2; }
  if (col >= xls->xls_colmax) { return -2; }
  if (row < xls->dim_rowmin) { xls->dim_rowmin = row; }
  if (row > xls->dim_rowmax) { xls->dim_rowmax = row; }
  if (col < xls->dim_colmin) { xls->dim_colmin = col; }
  if (col > xls->dim_colmax) { xls->dim_colmax = col; }

  xf = wsheet_xf(fmt);

  pkt = pkt_init(0, VARIABLE_PACKET);
  /* Write header */
  pkt_add16_le(pkt, name);
  pkt_add16_le(pkt, length);

  /* Write data */
  pkt_add16_le(pkt, row);
  pkt_add16_le(pkt, col);
  pkt_add16_le(pkt, xf);

  /* Write the number */
  memcpy(xl_double, &num, sizeof(xl_double));
  if (biff->byte_order)
    reverse(xl_double, sizeof(xl_double));

  pkt_addraw(pkt, xl_double, sizeof(xl_double));
  biff->append(biff, pkt->data, pkt->len);
  pkt_free(pkt);

  return 0;
}

/* Write a string to the specified row and column (zero indexed).
 * NOTE: There is an Excel 5 defined limit of 255 characters.
 * This writes the Excel LABEL record (BIFF3-BIFF5) */
int xls_writef_string(struct wsheetctx *xls, int row, int col, char *str, struct xl_format *fmt)
{
  uint16_t name = 0x0204; /* Record identifier */
  uint16_t length = 0x0008; /* Number of bytes to follow */
  uint16_t xf; /* The cell format */
  int len;
  struct pkt *pkt;

  len = strlen(str);

  /* LABEL must be < 255 chars */
  if (len > xls->xls_strmax) len = xls->xls_strmax;

  length += len;

  if (row >= xls->xls_rowmax) { return -2; }
  if (col >= xls->xls_colmax) { return -2; }
  if (row < xls->dim_rowmin) { xls->dim_rowmin = row; }
  if (row > xls->dim_rowmax) { xls->dim_rowmax = row; }
  if (col < xls->dim_colmin) { xls->dim_colmin = col; }
  if (col > xls->dim_colmax) { xls->dim_colmax = col; }

  xf = wsheet_xf(fmt);

  pkt = pkt_init(0, VARIABLE_PACKET);

  /* Write header */
  pkt_add16_le(pkt, name);
  pkt_add16_le(pkt, length);

  /* Write data */
  pkt_add16_le(pkt, row);
  pkt_add16_le(pkt, col);
  pkt_add16_le(pkt, xf);
  pkt_add16_le(pkt, len);

  pkt_addraw(pkt, (unsigned char *)str, len);
  wsheet_append(xls, pkt->data, pkt->len);
  pkt_free(pkt);
  return 0;
}

/* Write Worksheet BLANK record  (BIFF3-8) */
int xls_write_blank(struct wsheetctx *xls, int row, int col, struct xl_format *fmt)
{
  uint16_t name = 0x0201; /* Record identifier */
  uint16_t length = 0x0006; /* Number of bytes to follow */
  uint16_t xf; /* The cell format */
  struct pkt *pkt;

  if (row >= xls->xls_rowmax) { return -2; }
  if (col >= xls->xls_colmax) { return -2; }
  if (row < xls->dim_rowmin) { xls->dim_rowmin = row; }
  if (row > xls->dim_rowmax) { xls->dim_rowmax = row; }
  if (col < xls->dim_colmin) { xls->dim_colmin = col; }
  if (col > xls->dim_colmax) { xls->dim_colmax = col; }

  xf = wsheet_xf(fmt);

  pkt = pkt_init(10, FIXED_PACKET);

  /* Write header */
  pkt_add16_le(pkt, name);
  pkt_add16_le(pkt, length);

  /* Write data */
  pkt_add16_le(pkt, row);
  pkt_add16_le(pkt, col);
  pkt_add16_le(pkt, xf);

  wsheet_append(xls, pkt->data, pkt->len);
  pkt_free(pkt);
  return 0;
}

int xls_write_string(struct wsheetctx *xls, int row, int col, char *str)
{
  return xls_writef_string(xls, row, col, str, NULL);
}

int xls_write_number(struct wsheetctx *xls, int row, int col, double num)
{
  return xls_writef_number(xls, row, col, num, NULL);
}

void wsheet_set_column(struct wsheetctx *ws, int fcol, int lcol, int width)
{
  struct col_info *ci;

  /* First look and see if something like this already exists */
  TAILQ_FOREACH(ci, &ws->colinfos, cis) {
    if (ci->first_col == fcol && ci->last_col == lcol) {
      /* Found a match, just update the width */
      ci->col_width = width;
      return;
    }
  }

  ci = malloc(sizeof(struct col_info));
  ci->first_col = fcol;
  ci->last_col = lcol;
  ci->col_width = width;
  ci->xf = 0xF;
  ci->grbit = 0;

  TAILQ_INSERT_TAIL(&ws->colinfos, ci, cis);
}

void wsheet_store_selection(struct wsheetctx *xls, int frow, int fcol, int lrow, int lcol)
{
  struct pkt *pkt;
  int tmp;

  pkt = pkt_init(0, VARIABLE_PACKET);

  /* Swap rows and columns around */
  if (frow > lrow) {
    tmp = frow;
    frow = lrow;
    lrow = tmp;
  }

  if (fcol > lcol) {
    tmp = fcol;
    fcol = lcol;
    lcol = fcol;
  }

  /* Write header */
  pkt_add16_le(pkt, 0x001D);  /* Record identifier */
  pkt_add16_le(pkt, 0x000F);  /* Number of bytes to follow */

  pkt_add8(pkt, 3);  /* Pane position */
  pkt_add16_le(pkt, frow);  /* Active row */
  pkt_add16_le(pkt, fcol);  /* Active column */
  pkt_add16_le(pkt, 0);     /* Active cell ref */
  pkt_add16_le(pkt, 1);     /* Number of refs */
  pkt_add16_le(pkt, frow);  /* First row in reference */
  pkt_add16_le(pkt, lrow);  /* Last row in reference */
  pkt_add8(pkt, fcol);      /* First col in reference */
  pkt_add8(pkt, lcol);      /* Last col in reference */

  wsheet_append(xls, pkt->data, pkt->len);
  pkt_free(pkt);
}

/* Write BIFF record DEFCOLWIDTH if COLINFO records are in use. */
void wsheet_store_defcol(struct wsheetctx *wsheet)
{
  struct pkt *pkt;

  pkt = pkt_init(6, FIXED_PACKET);

  /* Write header */
  pkt_add16_le(pkt, 0x0055);  /* Record identifier */
  pkt_add16_le(pkt, 0x0002);  /* Number of bytes to follow */

  /* Write data */
  pkt_add16_le(pkt, 0x0008);  /* Default column width */

  bw_prepend((struct bwctx *)wsheet, pkt->data, pkt->len);
  pkt_free(pkt);
}

/* Write BIFF record COLINFO to define column widths
 *
 * Note: The SDK says the record length is 0x0B but Excel writes a 0x0C
 * length record */
void wsheet_store_colinfo(struct wsheetctx *wsheet, struct col_info *ci)
{
  struct pkt *pkt;
  float tmp;

  pkt = pkt_init(0, VARIABLE_PACKET);

  tmp = (float)ci->col_width + 0.72;   /* Fudge.  Excel subtracts 0.72 !? */
  tmp *= 256;  /* Convert to units of 1/256 of a char */

  /* Write header */
  pkt_add16_le(pkt, 0x007D);  /* Record identifier */
  pkt_add16_le(pkt, 0x000B);  /* Number of bytes to follow */

  /* Write data */
  pkt_add16_le(pkt, ci->first_col);  /* First formatted column */
  pkt_add16_le(pkt, ci->last_col);   /* Last formatted column */
  pkt_add16_le(pkt, (int)tmp);       /* Column width */
  pkt_add16_le(pkt, ci->xf);         /* XF */
  pkt_add16_le(pkt, ci->grbit);      /* Option flags */
  pkt_add8(pkt, 0x00);               /* Reserved */

  bw_prepend((struct bwctx *)wsheet, pkt->data, pkt->len);
  pkt_free(pkt);
}

/* write_url
 *
 * Write a hyperlink.  This is comprised of two elements: this visible label
 * and the invisible link.  The visible label is the same as the link unless an
 * alternative string is specified.  The label is written using the
 * write_string function.  Therefore the 255 character string limit applies.
 */
int wsheet_write_url(struct wsheetctx *wsheet, int row, int col, char *url, char *string, struct xl_format *fmt)
{
  struct pkt *pkt;
  int length;
  char *str;
  unsigned char unknown[40] =
  { 0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82,
    0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B, 0x02, 0x00, 0x00, 0x00,
    0x03, 0x00, 0x00, 0x00, 0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA,
    0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B
  };

  str = string;
  if (str == NULL)
    str = url;

  xls_writef_string(wsheet, row, col, str, fmt);

  length = 0x0034 + 2 * (1 + strlen(url));
  pkt = pkt_init(0, VARIABLE_PACKET);

  /* Write header */
  pkt_add16_le(pkt, 0x01B8);  /* Record identifier */
  pkt_add16_le(pkt, length);  /* Number of bytes to follow */

  pkt_add16_le(pkt, row);     /* Row number */
  pkt_add16_le(pkt, row);     /* Row number */
  pkt_add16_le(pkt, col);     /* Column number */
  pkt_add16_le(pkt, col);     /* Column number */
  pkt_addraw(pkt, unknown, sizeof(unknown));
  pkt_add32_le(pkt, strlen(url));

  pkt_addraw(pkt, (unsigned char *)url, strlen(url));
  wsheet_append(wsheet, pkt->data, pkt->len);
  pkt_free(pkt);

  return 0;
}

/* This method is used to set the height and XF format for a row.
 * Writes the BIFF record ROW. */
void wsheet_set_row(struct wsheetctx *wsheet, int row, int height, struct xl_format *fmt)
{
  struct pkt *pkt;
  int rowHeight;
  uint16_t xf; /* The cell format */

  rowHeight = height;
  /* Use -1 to set a format without setting height */
  if (rowHeight < 0)
    rowHeight = 0xff;
  else rowHeight *= 20;

  xf = wsheet_xf(fmt);

  pkt = pkt_init(0, VARIABLE_PACKET);

  /* Write header */
  pkt_add16_le(pkt, 0x0208);  /* Record identifier */
  pkt_add16_le(pkt, 0x0010);  /* Number of bytes to follow */

  /* Write payload */
  pkt_add16_le(pkt, row);     /* Row number */
  pkt_add16_le(pkt, 0x0000);  /* First defined column */
  pkt_add16_le(pkt, 0x0000);  /* Last defined column */
  pkt_add16_le(pkt, rowHeight); /* Row height */
  pkt_add16_le(pkt, 0x0000);  /* Used by Excel to optimise loading */
  pkt_add16_le(pkt, 0x0000);  /* Reserved */
  pkt_add16_le(pkt, 0x01C0);  /* Option flags. */
  pkt_add16_le(pkt, xf);      /* XF index */

  wsheet_append(wsheet, pkt->data, pkt->len);
  pkt_free(pkt);
}
