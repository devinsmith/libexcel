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

#ifndef __XLS_BIFFWRITER_H__
#define __XLS_BIFFWRITER_H__

#include <sys/types.h>
#include <stdint.h>
#include <stdio.h>

struct bwctx {
  int byte_order;
  unsigned char *data;
  unsigned int _sz;  /* Don't touch */
  unsigned int datasize;

  void (*append) (void *xlsctx, void *data, size_t size);
};

struct bwctx * bw_new(void);
void bw_destroy(struct bwctx *bw);
void bw_store_eof(struct bwctx *bw);
void bw_store_bof(struct bwctx *bw, uint16_t type);
void bw_append(void *xlsctx, void *data, size_t size);
void bw_prepend(struct bwctx *bw, void *data, size_t size);

void reverse(unsigned char *data, int size);

#endif /* __XLS_BIFFWRITER_H__ */
