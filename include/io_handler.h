/*
 * Copyright (c) 2014 Guilherme Steinmann <guidefloripa@gmail.com>
 *
 * Permission to use, copy, modify, and distribute this software for any
 * purpose with or without fee is hereby granted, provided that the above
 * copyright notice and this permission notice appear in all copies.
 *
 * THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
 * WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
 * MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BELIABLE FOR
 * ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
 * WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
 * ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
 * OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
 */

#ifndef __XLS_IO_HANDLER_H__
#define __XLS_IO_HANDLER_H__

#include <stdio.h>

struct xl_io_handler {
  void* (*create)(const char *filename);
  int   (*write )(void* handle, const void* buffer, size_t size);
  int   (*close )(void* handle);
};

void* xl_file_create(const char *filename);
int   xl_file_write(void *handle,const void* buffer,size_t size);
int   xl_file_close(void *handle);

extern struct xl_io_handler xl_file_handler;

#endif /* __XLS_IO_HANDLER_H__ */
