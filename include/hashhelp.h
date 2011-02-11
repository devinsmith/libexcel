/*
 * Copyright (c) 2010 Devin Smith <devin@devinsmith.net>
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

#ifndef __XLS_HASHHELP_H__
#define __XLS_HASHHELP_H__

/* This is not a generic hashtable implementation. The data type is a simple
 * integer (which was all I needed) */

struct hashnode {
  int key;
  int data;
  struct hashnode *next;
};

struct htbl {
  size_t filled;
  size_t size;
  struct hashnode **nodes;
};

struct htbl *hashtbl_new(size_t size);
void hashtbl_destroy(struct htbl *tbl);
int hashtbl_insert(struct htbl *tbl, int key, int data);
int hashtbl_get(struct htbl *tbl, int key);

#endif /* __XLS_HASHHELP_H__ */
