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

#ifndef __XLS_OLEWRITER_H__
#define __XLS_OLEWRITER_H__

#include <sys/types.h>
#include <stdint.h>
#include <stdio.h>
#include "io_handler.h"

struct owctx {
  const char *olefilename;
  struct xl_io_handler io_handler;
  void* io_handle;
  int fileclosed;
  int biff_only;
  int size_allowed;
  int biffsize;
  int booksize;
  int big_blocks;
  int list_blocks;
  int root_start;
  int block_count;
};

struct owctx * ow_new(const char *filename);
struct owctx * ow_new_ex(struct xl_io_handler io_handler, const char *filename);
void ow_destroy(struct owctx *ow);
int ow_set_size(struct owctx *ow, int biffsize);
void ow_write_header(struct owctx *ow);
void ow_write(struct owctx *ow, void *data, size_t size);
void ow_close(struct owctx *ow);

#endif /* __XLS_OLEWRITER_H__ */
