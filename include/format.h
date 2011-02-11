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

#ifndef __XLS_FORMAT_H__
#define __XLS_FORMAT_H__

#include <stdio.h>

#include "biffwriter.h"
#include "bsdqueue.h"
#include "stream.h"

struct xl_format {
  int xf_index;
  int font_index;
  char *fontname;
  int size;
  int bold;
  int italic;
  int color;
  int underline;
  int font_strikeout;
  int font_outline;
  int font_shadow;
  int font_script;
  int font_family;
  int font_charset;
  char *num_format_str;
  int num_format;
  int text_h_align;
  int text_wrap;
  int text_v_align;
  int text_justlast;
  int rotation;
  int fg_color;
  int bg_color;
  int pattern;
  int bottom;
  int top;
  int left;
  int right;
  int bottom_color;
  int top_color;
  int left_color;
  int right_color;
};

struct xl_format *fmt_new(int idx);
void fmt_destroy(struct xl_format *fmt);
struct pkt *fmt_get_font(struct xl_format *fmt);
struct pkt *fmt_get_xf(struct xl_format *fmt, int style);
int fmt_gethash(struct xl_format *fmt);

/* format setters */
void fmt_set_bold(struct xl_format *fmt, int bold_val);
void fmt_set_color(struct xl_format *fmt, char *colorname);
void fmt_set_align(struct xl_format *fmt, char *align);
void fmt_set_size(struct xl_format *fmt, int size);
void fmt_set_font(struct xl_format *fmt, char *font);
void fmt_set_colori(struct xl_format *fmt, int colorval);
void fmt_set_num_format(struct xl_format *fmt, int format);
void fmt_set_border_color(struct xl_format *fmt, char *colorname);
void fmt_set_border(struct xl_format *fmt, int format);
void fmt_set_pattern(struct xl_format *fmt, int pattern);
void fmt_set_bg_color(struct xl_format *fmt, char *colorname);
void fmt_set_fg_color(struct xl_format *fmt, char *colorname);
void fmt_set_text_wrap(struct xl_format *fmt, int val);
void fmt_set_rotation(struct xl_format *fmt, int val);
void fmt_set_merge(struct xl_format *fmt);
void fmt_set_underline(struct xl_format *fmt, int val);
void fmt_set_num_format_str(struct xl_format *fmt, char *str);

#endif /* __XLS_FORMAT_H__ */
