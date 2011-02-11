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

#include "olewriter.h"
#include "stream.h"

int ow_init(struct owctx *ow, char *filename);
void ow_write_pps(struct owctx *ow, char *name, int pps_type, int pps_dir, int pps_start, int pps_size);
void ow_write_property_storage(struct owctx *ow);
void ow_write_padding(struct owctx *ow);
void ow_write_big_block_depot(struct owctx *ow);

struct owctx * ow_new(char *filename)
{
  struct owctx *ow;

  ow = malloc(sizeof(struct owctx));

  if (ow_init(ow, filename) == -1) {
    free(ow);
    return NULL;
  }

  return ow;
}

void ow_destroy(struct owctx *ow)
{
  if (!ow->fileclosed)
    ow_close(ow);
  free(ow);
}

int ow_init(struct owctx *ow, char *filename)
{
  FILE *fp;

  ow->olefilename = filename;
  ow->filehandle = NULL;
  ow->fileclosed = 0;
  ow->biff_only = 0;
  ow->size_allowed = 0;
  ow->biffsize = 0;
  ow->booksize = 0;
  ow->big_blocks = 0;
  ow->list_blocks = 0;
  ow->root_start = 0;
  ow->block_count = 4;

  if (filename == NULL)
    return -1;

  /* Open file for writing */
  fp = fopen(filename, "wb");
  if (fp == NULL)
    return -1;

  ow->filehandle = fp;

  return 0;
}

/****************************************************************************
 * ow_set_size(struct owctx *ow, int biffsize)
 *
 * Set the size of the data to be written to the OLE stream
 *
 *  big_blocks = (109 depot block * (128 -1 marker word) -
 *               (1 * end words)) = 13842
 *  maxsize = big_blocks * 512 bytes = 7087104
 */
int ow_set_size(struct owctx *ow, int biffsize)
{
  int maxsize = 7087104; /* TODO: extend max size */

  if (biffsize > maxsize) {
    ow->size_allowed = 0;
    return 0;
  }

  ow->biffsize = biffsize;
  /* Set the min file size to 4k to avoid having to use small blocks */
  if (biffsize > 4096) {
    ow->booksize = biffsize;
  } else {
    ow->booksize = 4096;
  }

  ow->size_allowed = 1;
  return 1;
}

/****************************************************************************
 * ow_calculate_sizes(struct owctx *ow)
 *
 * Calculate various sizes need for the OLE stream
 */
void ow_calculate_sizes(struct owctx *ow)
{
  int datasize = ow->booksize;

  if (datasize % 512 == 0) {
    ow->big_blocks = datasize / 512;
  } else {
    ow->big_blocks = (datasize / 512) + 1;
  }

  /* There are 127 list blocks and 1 marker blocks for each big block
   * depot + 1 end of chain block */
  ow->list_blocks = (ow->big_blocks / 127) + 1;
  ow->root_start = ow->big_blocks;
}

/****************************************************************************
 * ow_write_header(struct owctx *ow)
 *
 * Write OLE header block.
 */
void ow_write_header(struct owctx *ow)
{
  int root_start;
  int num_lists;
  struct pkt *pkt;
  int i;

  if (ow->biff_only)
    return;

  ow_calculate_sizes(ow);

  root_start = ow->root_start;
  num_lists = ow->list_blocks;

  pkt = pkt_init(0, VARIABLE_PACKET);
  pkt_add32(pkt, 0xD0CF11E0); /* OLE document file id part 1 */
  pkt_add32(pkt, 0xA1B11AE1); /* OLE document file id part 2 */
  pkt_add32_le(pkt, 0x00); /* UID of this file (can be all 0's) 1/4 */
  pkt_add32_le(pkt, 0x00); /* UID of this file (can be all 0's) 2/4 */
  pkt_add32_le(pkt, 0x00); /* UID of this file (can be all 0's) 3/4 */
  pkt_add32_le(pkt, 0x00); /* UID of this file (can be all 0's) 4/4 */
  pkt_add16_le(pkt, 0x3E); /* Revision number (almost always 0x003E) */
  pkt_add16_le(pkt, 0x03); /* Version number (almost always 0x0003) */
  pkt_add16(pkt, 0xFEFF);  /* Byte order identifier:
                            * (0xFEFF = Little Endian)
                            * (0xFFFE = Big Endian)    */
  pkt_add16_le(pkt, 0x09); /* 2^x  (9 = 512 bytes) */
  pkt_add32_le(pkt, 0x06); /* Unknown 5 */
  pkt_add32_le(pkt, 0x00); /* Unknown 5 */
  pkt_add32_le(pkt, 0x00); /* Unknown 5 */
  pkt_add32_le(pkt, num_lists); /* num_bbd_blocks */
  pkt_add32_le(pkt, root_start); /* root_startblock */
  pkt_add32_le(pkt, 0x00); /* Unknown 6 */
  pkt_add32_le(pkt, 0x1000); /* Unknown 6 */
  pkt_add32_le(pkt, -2);  /* sbd_startblock */
  pkt_add32_le(pkt, 0x00); /* Unknown 7 */
  pkt_add32_le(pkt, -2); /* Unknown 7 */
  pkt_add32_le(pkt, 0x00); /* Unknown 7 */

  for (i = 1; i <= num_lists; i++) {
    root_start++;
    pkt_add32_le(pkt, root_start);

  }

  for (i = num_lists; i <= 108; i++) {
    pkt_add32_le(pkt, -1); /* Unused */
  }

  fwrite(pkt->data, 1, pkt->len, ow->filehandle);
  pkt_free(pkt);
}

/****************************************************************************
 * ow_close(struct owctx *ow)
 *
 * Write root entry, big block list and close the filehandle.
 * This routine is used to explicity close the open filehandle without
 * having to wait for destroy.
 */
void ow_close(struct owctx *ow)
{
  if (!ow->size_allowed)
    return;

  if (!ow->biff_only) {
    ow_write_padding(ow);
    ow_write_property_storage(ow);
    ow_write_big_block_depot(ow);
  }
  fclose(ow->filehandle);
  ow->fileclosed = 1;
}

/****************************************************************************
 * ow_write(struct owctx *ow)
 *
 * Write BIFF data to OLE file
 */
void ow_write(struct owctx *ow, void *data, size_t len)
{
  fwrite(data, 1, len, ow->filehandle);
}

/****************************************************************************
 * ow_write_big_block_depot(struct owctx *ow)
 *
 * Write big block depot.
 */
void ow_write_big_block_depot(struct owctx *ow)
{
  int num_blocks = ow->big_blocks;
  int num_lists = ow->list_blocks;
  int total_blocks = num_lists * 128;
  int used_blocks = num_blocks + num_lists + 2;
  struct pkt *pkt;
  int i;

  pkt = pkt_init(0, VARIABLE_PACKET);
  for (i = 1; i <= num_blocks - 1; i++) {
    pkt_add32_le(pkt, i);
  }

  /* End of chain */
  pkt_add32_le(pkt, -2);
  pkt_add32_le(pkt, -2);

  for (i = 1; i <= num_lists; i++) {
    pkt_add32_le(pkt, -3);
  }

  for (i = used_blocks; i <= total_blocks; i++) {
    pkt_add32_le(pkt, -1);
  }

  fwrite(pkt->data, 1, pkt->len, ow->filehandle);
  pkt_free(pkt);
}

/****************************************************************************
 * ow_write_property_storage(struct owctx *ow)
 *
 * Write property storage.  TODO: add summary sheets
 */
void ow_write_property_storage(struct owctx *ow)
{
  //int rootsize = -2;
  int booksize = ow->booksize;

  ow_write_pps(ow, "Root Entry", 0x05, 1, -2, 0x00);
  ow_write_pps(ow, "Book", 0x02, -1, 0x00, booksize);
  ow_write_pps(ow, NULL, 0x00, -1, 0x00, 0x0000);
  ow_write_pps(ow, NULL, 0x00, -1, 0x00, 0x0000);
}

/****************************************************************************
 * ow_write_pps(struct owctx *ow, char *name)
 *
 * Write property sheet in property storage
 */
void ow_write_pps(struct owctx *ow, char *name, int pps_type, int pps_dir, int pps_start, int pps_size)
{
  unsigned char header[64];
  int length;
  struct pkt *pkt;

  memset(header, 0, sizeof(header));
  length = 0;
  if (name != NULL)
  {
    /* Simulate a unicode string */
    char *p = name;
    int i = 0;
    while (*p != '\0') {
      header[i] = *p++;
      i += 2;
    }
    length = (strlen(name) * 2) + 2;
  }

  pkt = pkt_init(0, VARIABLE_PACKET);
  pkt_addraw(pkt, header, sizeof(header));
  pkt_add16_le(pkt, length); /* pps_sizeofname 0x40 */
  pkt_add16_le(pkt, pps_type); /* 0x42 */
  pkt_add32_le(pkt, -1); /* pps_prev  0x44 */
  pkt_add32_le(pkt, -1); /* pps_next  0x48 */
  pkt_add32_le(pkt, pps_dir); /* pps_dir   0x4C */

  pkt_add32_le(pkt, 0);  /* unknown   0x50 */
  pkt_add32_le(pkt, 0);  /* unknown   0x54 */
  pkt_add32_le(pkt, 0);  /* unknown   0x58 */
  pkt_add32_le(pkt, 0);  /* unknown   0x5C */
  pkt_add32_le(pkt, 0);  /* unknown   0x60 */

  pkt_add32_le(pkt, 0);  /* pps_ts1s  0x64 */
  pkt_add32_le(pkt, 0);  /* pps_ts1d  0x68 */
  pkt_add32_le(pkt, 0);  /* pps_ts2s  0x6C */
  pkt_add32_le(pkt, 0);  /* pps_ts2d  0x70 */
  pkt_add32_le(pkt, pps_start); /* pps_start 0x74 */
  pkt_add32_le(pkt, pps_size); /* pps_size 0x78 */
  pkt_add32_le(pkt, 0);  /* unknown  0x7C */

  fwrite(pkt->data, 1, pkt->len, ow->filehandle);
  pkt_free(pkt);
}

/****************************************************************************
 * ow_write_padding(struct owctx *ow)
 *
 * Pad the end of the file
 */
void ow_write_padding(struct owctx *ow)
{
  int biffsize = ow->biffsize;
  int min_size;

  if (biffsize < 4096) {
    min_size = 4096;
  } else {
    min_size = 512;
  }

  if (biffsize % min_size != 0) {
    int padding = min_size - (biffsize % min_size);
    unsigned char *buffer;

    buffer = malloc(padding);
    memset(buffer, 0, padding);
    fwrite(buffer, 1, padding, ow->filehandle);
    free(buffer);
  }
}

