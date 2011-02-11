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

#include "biffwriter.h"
#include "stream.h"

const int g_BIFF_version = 0x0500;

int bw_init(struct bwctx *bw);

struct bwctx * bw_new(void)
{
  struct bwctx *bw;

  bw = malloc(sizeof(struct bwctx));

  if (bw_init(bw) == -1) {
    free(bw);
    return NULL;
  }

  return bw;
}

void bw_destroy(struct bwctx *bw)
{
  free(bw->data);
  free(bw);
}

void reverse(unsigned char *data, int size)
{
  unsigned char *t, *m;

  t = data + size;
  m = data + size / 2;
  while (data != m) {
    unsigned char c = *data;
    *data++ = *--t;
    *t = c;
  }
}

int bw_setbyteorder(struct bwctx *bw)
{
  double d = 1.2345;
  unsigned char data[8];
  unsigned char hexdata[8] = {0x8D, 0x97, 0x6E, 0x12, 0x83, 0xC0, 0xF3, 0x3F};

  /* Test for endian */
  memcpy(data, &d, sizeof(data));
  if (memcmp(data, hexdata, sizeof(data))) {
    bw->byte_order = 1;
  } else {
    reverse(data, sizeof(data));
    if (memcmp(data, hexdata, sizeof(data))) {
      bw->byte_order = 0;
    } else {
      return -1;
    }
  }
  return 0;
}

int bw_init(struct bwctx *bw)
{
  bw->data = NULL;
  bw->datasize = 0;
  bw->_sz = 0;
  bw_setbyteorder(bw);
  bw->append = bw_append;

  return 0;
}

#define ROUNDVAL   16
#define ROUNDUP(n)  (((n)+ROUNDVAL-1)&-ROUNDVAL)

void bw_resize(struct bwctx *bw, size_t size)
{
  if (bw->_sz != size) {
    if (size > 0) {
      if (bw->data == NULL)
        bw->data = malloc(ROUNDUP(1 + size));
      else
        bw->data = realloc(bw->data, ROUNDUP(1 + size));
      bw->_sz = size;
    } else if (bw->data != NULL) {
      free(bw->data);
      bw->data = NULL;
    }
  }
}

void bw_append(void *xlsctx, void *data, size_t size)
{
  struct bwctx *bw = (struct bwctx *)xlsctx;

  int len = bw->_sz;
  bw_resize(bw, len + size);
  memcpy(bw->data + len, data, size);
  bw->datasize += size;
}

void bw_prepend(struct bwctx *bw, void *data, size_t size)
{
  int len = bw->_sz;
  bw_resize(bw, len + size);
  memmove(bw->data + size, bw->data, len);
  memcpy(bw->data, data, size);
  bw->datasize += size;
}


/****************************************************************************
 * bw_store_bof(struct bwctx *bw, uint16_t type)
 *
 * type = 0x0005, Workbook
 * type = 0x0010, Worksheet
 *
 * Writes Excel BOF (Beginning of File) record to indicate the beginning of
 * a stream or substream in the BIFF file
 */
void bw_store_bof(struct bwctx *bw, uint16_t type)
{
  uint16_t name = 0x0809;   /* Record identifier */
  uint16_t length = 0x0008; /* Number of bytes to follow */

  /* According to the SDK "build" and "year" should be set to zero.
   * However, this throws a warning in Excel 5.  So, use these
   * magic umbers */
  uint16_t build = 0x096C;
  uint16_t year = 0x07C9;

  struct pkt *pkt;

  pkt = pkt_init(12, FIXED_PACKET);

  /* Construct header */
  pkt_add16_le(pkt, name);
  pkt_add16_le(pkt, length);

  /* Construct data */
  pkt_add16_le(pkt, g_BIFF_version);
  pkt_add16_le(pkt, type);
  pkt_add16_le(pkt, build);
  pkt_add16_le(pkt, year);
  bw_prepend(bw, pkt->data, pkt->len);
  pkt_free(pkt);
}

/****************************************************************************
 * bw_store_eof(struct bwctx *bw)
 *
 * Writes Excel EOF (End of File) record to indicate the end of a BIFF stream.
 */
void bw_store_eof(struct bwctx *bw)
{
  struct pkt *pkt;
  uint16_t name = 0x000A;   /* Record identifier */
  uint16_t length = 0x0000; /* Number of bytes to follow */

  pkt = pkt_init(4, FIXED_PACKET);

  /* Construct header */
  pkt_add16_le(pkt, name);
  pkt_add16_le(pkt, length);

  bw->append(bw, pkt->data, pkt->len);
  pkt_free(pkt);
}
