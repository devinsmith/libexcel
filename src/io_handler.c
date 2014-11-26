/*
 * Copyright (c) 2014 Guilherme Steinmann <guidefloripa@gmail.com>
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

#include "io_handler.h"

struct xl_io_handler xl_file_handler = {
	xl_file_create,
	xl_file_write,
	xl_file_close
};

void* xl_file_create(const char *filename)
{
	return filename ? (void*)fopen(filename,"wb") : NULL;
}

int xl_file_write(void *handle,const void* buffer,size_t size)
{
	return handle ? fwrite(buffer,1,size,(FILE*)handle) : -1;
}

int xl_file_close(void *handle)
{
	return handle ? fclose((FILE*)handle) : -1;
}
