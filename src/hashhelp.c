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

#include <stdlib.h>
#include "hashhelp.h"

struct htbl *hashtbl_new(size_t size)
{
  struct htbl *table;

  table = malloc(sizeof(struct htbl));
  if (table == NULL)
    return NULL;

  table->nodes = calloc(size, sizeof(struct hashnode *));
  if (table->nodes == NULL) {
    free(table);
    return NULL;
  }

  if (size <= 0)
    size = 10;  /* A sane default */

  table->filled = 0;
  table->size = size;
  return table;
}

void hashtbl_destroy(struct htbl *tbl)
{
  size_t n;
  struct hashnode *node, *oldnode;

  for (n = 0; n < tbl->size; n++) {
    node = tbl->nodes[n];
    while (node) {
      oldnode = node;
      node = node->next;
      free(oldnode);
    }
  }
  free(tbl->nodes);
  free(tbl);
}

static void hashtbl_resize(struct htbl *tbl, int new_size)
{
  struct htbl *newtbl;
  size_t n;
  struct hashnode *node;

  newtbl = hashtbl_new(new_size);
  /* For some reason we can't allocate new nodes, but that's actually
   * ok because we'll just use the old inefficient hash table. */
  if (newtbl == NULL)
    return;

  for (n = 0; n < tbl->size; n++) {
    node = tbl->nodes[n];
    while (node) {
      hashtbl_insert(newtbl, node->key, node->data);
      node = node->next;
    }
  }
  free(tbl->nodes);
  tbl->nodes = newtbl->nodes;
  tbl->filled = newtbl->filled;
  tbl->size = newtbl->size;
  free(newtbl);
}

int hashtbl_insert(struct htbl *tbl, int key, int data)
{
  struct hashnode *node;
  size_t hash = key % tbl->size;

  if (tbl->filled >= (tbl->size << 2)) {
    hashtbl_resize(tbl, (tbl->size << 2));
  }

  node = tbl->nodes[hash];
  while (node) {
    if (node->key == key) {
      node->data = data;
      return 0;
    }
    node = node->next;
  }

  if (!(node = malloc(sizeof(struct hashnode))))
    return -1;
  node->key = key;
  node->data = data;
  node->next = tbl->nodes[hash];
  tbl->nodes[hash] = node;
  tbl->filled++;
  return 0;
}

int hashtbl_get(struct htbl *tbl, int key)
{
  struct hashnode *node;
  size_t hash = key % tbl->size;

  node = tbl->nodes[hash];
  while (node) {
    if (node->key == key) return node->data;
    node = node->next;
  }
  return -1;
}
