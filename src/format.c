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

#include "format.h"
#include "stream.h"

static int fmt_get_color(char *colorname);

struct xl_format * fmt_new(int idx)
{
  struct xl_format *ret;

  ret = malloc(sizeof(struct xl_format));

  ret->xf_index = idx;
  ret->font_index = 0;
  ret->fontname = strdup("Arial");
  ret->size = 10;
  ret->bold = 0x0190;
  ret->italic = 0;
  ret->color = 0x7FFF;
  ret->underline = 0;
  ret->font_strikeout = 0;
  ret->font_outline = 0;
  ret->font_shadow = 0;
  ret->font_script = 0;
  ret->font_family = 0;
  ret->font_charset = 0;

  ret->num_format = 0;
  ret->num_format_str = NULL;

  ret->text_h_align = 0;
  ret->text_wrap = 0;
  ret->text_v_align = 2;
  ret->text_justlast = 0;
  ret->rotation = 0;

  ret->fg_color = 0x40;
  ret->bg_color = 0x41;

  ret->pattern = 0;

  ret->bottom = 0;
  ret->top = 0;
  ret->left = 0;
  ret->right = 0;

  ret->bottom_color = 0x40;
  ret->top_color = 0x40;
  ret->left_color = 0x40;
  ret->right_color = 0x40;

  return ret;
}

void fmt_destroy(struct xl_format *fmt)
{
  free(fmt->fontname);
  free(fmt->num_format_str);
  free(fmt);
}

/*
 * Generate an Excel BIFF XF record. */
struct pkt *fmt_get_xf(struct xl_format *fmt, int style)
{
  struct pkt *pkt;
  unsigned int align;
  int atr_num;
  int atr_fnt;
  int atr_alc;
  int atr_bdr;
  int atr_pat;
  int atr_prot;
  unsigned int icv;
  unsigned int fill;
  unsigned int border1;
  unsigned int border2;

  pkt = pkt_init(20, FIXED_PACKET);
  if (pkt == NULL)
    return NULL;

  /* Flags to indicate if attributes have been set */
  atr_num = (fmt->num_format != 0);
  atr_fnt = (fmt->font_index != 0);
  atr_alc = fmt->text_wrap;
  atr_bdr = (fmt->bottom || fmt->top || fmt->left || fmt->right);
  atr_pat = (fmt->fg_color || fmt->bg_color || fmt->pattern);
  atr_prot = 0;

  /* Zero the default border colour if the border has not been set */
  if (fmt->bottom == 0) fmt->bottom_color = 0;
  if (fmt->top == 0) fmt->top_color = 0;
  if (fmt->right == 0) fmt->right_color = 0;
  if (fmt->left == 0) fmt->left_color = 0;

  align  = fmt->text_h_align;
  align |= fmt->text_wrap << 3;
  align |= fmt->text_v_align << 4;
  align |= fmt->text_justlast << 7;
  align |= fmt->rotation << 8;
  align |= atr_num << 10;
  align |= atr_fnt << 11;
  align |= atr_alc << 12;
  align |= atr_bdr << 13;
  align |= atr_pat << 14;
  align |= atr_prot << 15;

  icv  = fmt->fg_color;
  icv |= fmt->bg_color << 7;

  fill  = fmt->pattern;
  fill |= fmt->bottom << 6;
  fill |= fmt->bottom_color << 9;

  border1  = fmt->top;
  border1 |= fmt->left << 3;
  border1 |= fmt->right << 6;
  border1 |= fmt->top_color << 9;

  border2  = fmt->left_color;
  border2 |= fmt->right_color << 7;

  /* Write header */
  pkt_add16_le(pkt, 0x00E0);  /* Record identifier */
  pkt_add16_le(pkt, 0x0010);  /* Number of bytes to follow */

  /* Payload */
  pkt_add16_le(pkt, fmt->font_index); /* Index to FONT record */
  pkt_add16_le(pkt, fmt->num_format); /* Index to FORMAT record */
  pkt_add16_le(pkt, style);           /* Style and other options */
  pkt_add16_le(pkt, align);           /* Alignment */
  pkt_add16_le(pkt, icv);             /* Color palette and other options */
  pkt_add16_le(pkt, fill);            /* Fill and border line style */
  pkt_add16_le(pkt, border1);         /* Border line style and color */
  pkt_add16_le(pkt, border2);         /* Border color */

  return pkt;
}

/*
 * Generate an Excel BIFF FONT record. */
struct pkt *fmt_get_font(struct xl_format *fmt)
{
  struct pkt *pkt;
  int cch;
  int grbit;

  pkt = pkt_init(0, VARIABLE_PACKET);
  if (pkt == NULL)
    return NULL;

  cch = strlen(fmt->fontname);

  grbit = 0x00;
  if (fmt->italic)
    grbit |= 0x02;
  if (fmt->font_strikeout)
    grbit |= 0x08;
  if (fmt->font_outline)
    grbit |= 0x10;
  if (fmt->font_shadow)
    grbit |= 0x20;

  /* Write header */
  pkt_add16_le(pkt, 0x0031);  /* Record identifier */
  pkt_add16_le(pkt, 0x000F + cch);  /* Number of bytes to follow */

  /* Write data */
  pkt_add16_le(pkt, fmt->size * 20);   /* Height of font (1/20th of a point) */
  pkt_add16_le(pkt, grbit);            /* Font attributes */
  pkt_add16_le(pkt, fmt->color);       /* Index to color palette */
  pkt_add16_le(pkt, fmt->bold);        /* Bold style */
  pkt_add16_le(pkt, fmt->font_script); /* Superscript/subscript */
  pkt_add8(pkt, fmt->underline);       /* Underline */
  pkt_add8(pkt, fmt->font_family);     /* Font family */
  pkt_add8(pkt, fmt->font_charset);    /* Font charset */
  pkt_add8(pkt, 0x00);                 /* Reserved */
  pkt_add8(pkt, cch);                  /* Length of font name (count of ch) */
  pkt_addraw(pkt, (unsigned char *)fmt->fontname, cch);  /* Fontname */

  return pkt;
}

void fmt_set_bold(struct xl_format *fmt, int bold_val)
{
  if (bold_val > 0)
    fmt->bold = 0x2BC;
  else
    fmt->bold = 0x190;
}

struct key_value {
  char *name;
  int value;
};

void fmt_set_align(struct xl_format *fmt, char *align)
{
  int i;
  int num_vals;
  struct key_value halign[] = {
    {"left", 1},
    {"centre", 2},
    {"center", 2},
    {"right", 3},
    {"fill", 4},
    {"justify", 5},
    {"merge", 6}
  };
  struct key_value valign[] = {
    {"top", 0},
    {"vcentre", 1},
    {"vcenter", 1},
    {"bottom", 2},
    {"vjustify", 3}
  };

  num_vals = sizeof(halign) / sizeof(struct key_value);
  for (i = 0; i < num_vals; i++) {
    if (strcmp(align, halign[i].name) == 0) {
      fmt->text_h_align = halign[i].value;
      return;
    }
  }

  num_vals = sizeof(valign) / sizeof(struct key_value);
  for (i = 0; i < num_vals; i++) {
    if (strcmp(align, valign[i].name) == 0) {
      fmt->text_v_align = valign[i].value;
      return;
    }
  }
}

void fmt_set_merge(struct xl_format *fmt)
{
  fmt->text_h_align = 6;
}

void fmt_set_color(struct xl_format *fmt, char *colorname)
{
  fmt->color = fmt_get_color(colorname);
}

void fmt_set_text_wrap(struct xl_format *fmt, int val)
{
  fmt->text_wrap = val;
}

void fmt_set_border_color(struct xl_format *fmt, char *colorname)
{
  int color;

  color = fmt_get_color(colorname);
  fmt->bottom_color = color;
  fmt->top_color = color;
  fmt->left_color = color;
  fmt->right_color = color;
}

void fmt_set_bg_color(struct xl_format *fmt, char *colorname)
{
  fmt->bg_color = fmt_get_color(colorname);
}

void fmt_set_fg_color(struct xl_format *fmt, char *colorname)
{
  fmt->fg_color = fmt_get_color(colorname);
}

void fmt_set_border(struct xl_format *fmt, int style)
{
  fmt->bottom = style;
  fmt->top = style;
  fmt->left = style;
  fmt->right = style;
}

void fmt_set_colori(struct xl_format *fmt, int colorval)
{
  if (colorval < 8 || colorval > 63)
    fmt->color = 0x7FFF;

  fmt->color = colorval;
}

void fmt_set_pattern(struct xl_format *fmt, int pattern)
{
  fmt->pattern = pattern;
}

void fmt_set_rotation(struct xl_format *fmt, int val)
{
  fmt->rotation = val;
}

void fmt_set_size(struct xl_format *fmt, int size)
{
  fmt->size = size;
}

void fmt_set_num_format(struct xl_format *fmt, int format)
{
  fmt->num_format = format;
}

void fmt_set_font(struct xl_format *fmt, char *font)
{
  if (fmt->fontname)
    free(fmt->fontname);
  fmt->fontname = strdup(font);
}

void fmt_set_underline(struct xl_format *fmt, int val)
{
  fmt->underline = val;
}

void fmt_set_num_format_str(struct xl_format *fmt, char *str)
{
  fmt->num_format_str = strdup(str);
}

static int fmt_get_color(char *colorname)
{
  int num_colors;
  int i;

  struct key_value colors[] = {
      {"aqua", 0x0F},
      {"black", 0x08},
      {"blue", 0x0C},
      {"fuchsia", 0x0E},
      {"gray", 0x17},
      {"grey", 0x17},
      {"green", 0x11},
      {"lime", 0x0B},
      {"navy", 0x12},
      {"orange", 0x1D},
      {"purple", 0x24},
      {"red", 0x0A},
      {"silver", 0x16},
      {"white", 0x09},
      {"yellow", 0x0D}
  };

  num_colors = sizeof(colors) / sizeof(struct key_value);
  for (i = 0; i < num_colors; i++) {
    if (strcmp(colorname, colors[i].name) == 0)
      return colors[i].value;
  }

  return 0x7FFF;
}

static int fhc(char *str)
{
  int hash = 0;
  while (*str)
    hash = (31 * hash) + *str++;
  return hash;
}

int fmt_gethash(struct xl_format *fmt)
{
  int hash;

  hash = fhc(fmt->fontname);
  hash += fmt->size;
  hash += (fmt->font_script + fmt->underline);
  hash += (fmt->font_strikeout + fmt->bold + fmt->font_outline);
  hash += (fmt->font_family + fmt->font_charset);
  hash += (fmt->font_shadow + fmt->color + fmt->italic);
  if (fmt->num_format_str != NULL)
    hash += fhc(fmt->num_format_str);

  return hash;
}
