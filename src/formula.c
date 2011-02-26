/*
 * Copyright (c) 2011 Devin Smith <devin@devinsmith.net>
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

#include <sys/queue.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#include "formula.h"
#include "stream.h"

enum token_types {
  TOKEN_EQUALS,
  TOKEN_LPAREN,
  TOKEN_RPAREN,
  TOKEN_OPERATOR,
  TOKEN_CELL,
  TOKEN_CELLRANGE,
  TOKEN_FUNCTION,
  TOKEN_NUMBER,
  TOKEN_STRING,
  TOKEN_COMMA
};

enum tokenizer_state {
  TS_DEFAULT,
  TS_WORD,
  TS_NUMBER,
  TS_STRING
};

struct token {
  int type;
  char *data;
  TAILQ_ENTRY(token) tokens;
};

/* Declaration type */
TAILQ_HEAD(token_list, token);

#ifdef FORMULA_DEBUG
static void dump_hex(void *vp, int length);
#endif

int is_func(char *func)
{
  int num_funcs;
  int i;
  char *funcs[] = { "AVERAGE", "IF", "SIN", "SUM", "PI"
  };

  num_funcs = sizeof(funcs) / sizeof(funcs[0]);
  for (i = 0; i < num_funcs; i++) {
    if (strcmp(func, funcs[i]) == 0)
      return 1;
  }
  return 0;
}

void tokenize(char *inp, struct token_list *tl)
{
  char token[128];
  int tp;
  char ch;
  char *strpos = inp, *strend = inp + strlen(inp);
  int state;
  struct token *tn;

  state = TS_DEFAULT;
  tp = 0;
  tn = NULL;

  /* We can skip the first = */
  if (*strpos == '=')
    strpos++;

  while(strpos <= strend) {
    ch = *strpos;
    switch (state) {
    case TS_DEFAULT:
      if (ch == '\0') {
        break;
      }
      if (ch == '=') {
        tn = malloc(sizeof(struct token));
        tn->type = TOKEN_EQUALS;
        tn->data = NULL;
        TAILQ_INSERT_TAIL(tl, tn, tokens);
      } else if (ch == ' ') {
        /* Do nothing */
      } else if (ch == '(') {
        tn = malloc(sizeof(struct token));
        tn->type = TOKEN_LPAREN;
        tn->data = NULL;
        TAILQ_INSERT_TAIL(tl, tn, tokens);
      } else if (ch == ')') {
        tn = malloc(sizeof(struct token));
        tn->type = TOKEN_RPAREN;
        tn->data = NULL;
        TAILQ_INSERT_TAIL(tl, tn, tokens);
      } else if (ch == '+' || ch == '-' || ch == '>' || ch == '<' ||
          ch == '*' || ch == '/') {
        /* Special case, check if this is a negative number / cell */
        if (ch == '-' && tn != NULL && tn->type == TOKEN_OPERATOR) {
          token[tp++] = ch;
        } else {
          tn = malloc(sizeof(struct token));
          tn->type = TOKEN_OPERATOR;
          token[0] = ch;
          token[1] = '\0';
          tn->data = strdup(token);
          TAILQ_INSERT_TAIL(tl, tn, tokens);
        }
      } else if (ch == ',') {
        tn = malloc(sizeof(struct token));
        tn->type = TOKEN_COMMA;
        tn->data = NULL;
        TAILQ_INSERT_TAIL(tl, tn, tokens);
      } else if (ch >= 'A' && ch <= 'Z') {
        token[tp++] = ch;
        state = TS_WORD;
      } else if (ch >= '0' && ch <= '9') {
        token[tp++] = ch;
        state = TS_NUMBER;
      } else if (ch == '$') {
        token[tp++] = ch;
        state = TS_WORD;
      } else if (ch == '"') {
        state = TS_STRING;
      }
      else printf("UNKNOWN TOKEN\n");
      break;
    case TS_WORD:
      if (ch >= 'A' && ch <= 'Z') {
        token[tp++] = ch;
      } else if (ch == '$') {
        token[tp++] = ch;
      } else if (ch == ':') {
        token[tp++] = ch;
      } else if (ch >= '0' && ch <= '9') {
        token[tp++] = ch;
      } else {
        token[tp] = '\0';
        tp = 0;
        if (is_func(token)) {
          tn = malloc(sizeof(struct token));
          tn->type = TOKEN_FUNCTION;
          tn->data = strdup(token);
          TAILQ_INSERT_TAIL(tl, tn, tokens);
        } else if (strchr(token, ':') != NULL) {
          tn = malloc(sizeof(struct token));
          tn->type = TOKEN_CELLRANGE;
          tn->data = strdup(token);
          TAILQ_INSERT_TAIL(tl, tn, tokens);
        } else {
          /* Assume it's a cell? */
          tn = malloc(sizeof(struct token));
          tn->type = TOKEN_CELL;
          tn->data = strdup(token);
          TAILQ_INSERT_TAIL(tl, tn, tokens);
        }
        state = TS_DEFAULT;
        continue;
      }
      break;
    case TS_STRING:
      if (ch == '"') {
        token[tp] = '\0';
        tp = 0;
        tn = malloc(sizeof(struct token));
        tn->type = TOKEN_STRING;
        tn->data = strdup(token);
        TAILQ_INSERT_TAIL(tl, tn, tokens);
        state = TS_DEFAULT;
      } else
        token[tp++] = ch;
      break;
    case TS_NUMBER:
      if (ch < '0' || ch > '9') {
        token[tp] = '\0';
        tp = 0;
        tn = malloc(sizeof(struct token));
        tn->type = TOKEN_NUMBER;
        tn->data = strdup(token);
        TAILQ_INSERT_TAIL(tl, tn, tokens);
        state = TS_DEFAULT;
        continue;
      } else
        token[tp++] = ch;
      break;
    }
    strpos++;
  }
}

/* operators
 * precedence   operators       associativity
 * 1            !               right to left
 * 2            * / %           left to right
 * 3            + -             left to right
 * 4            =                right to left */
int op_preced(const char c)
{
  switch(c) {
    case '!':
      return 4;
    case '*':  case '/': case '%':
      return 3;
    case '+': case '-':
      return 2;
    case '=':
      return 1;
  }
  return 0;
}

int op_left_assoc(const char c)
{
  switch(c) {
    /* left to right */
    case '*': case '/': case '%': case '+': case '-':
      return 1;
    /* right to left */
    case '=': case '!':
      return 0;
  }
  return 0;
}

void encode_operator(struct pkt *pkt, const char op)
{
  switch (op) {
  case '+':
    pkt_add8(pkt, 0x03); /* tAdd */
    break;
  case '*':
    pkt_add8(pkt, 0x05); /* tMul */
  }
}

void encode_number(struct pkt *pkt, const char *data)
{
  int number;

  number = strtol(data, NULL, 10);
  pkt_add8(pkt, 0x1E); /* tInt */
  if (number < 0) {
    pkt_add16_le(pkt, -number);
    pkt_add8(pkt, 0x13); /* tUminus */
  } else
    pkt_add16_le(pkt, number);
}

/* Based on code from:
 * http://en.wikipedia.org/wiki/Shunting-yard_algorithm */
int parse_token_list(struct token_list *tlist, struct pkt *pkt)
{
  struct token *token;
  struct token *stack[32];
  unsigned int sl;  /* Stack length */
  struct token *sctoken;

  sl = 0;
  TAILQ_FOREACH(token, tlist, tokens) {
    if (token->type == TOKEN_NUMBER) {
      encode_number(pkt, token->data);
    } else if (token->type == TOKEN_OPERATOR) {
      while (sl > 0) {
        sctoken = stack[sl - 1];
        /* While there is an operator token, o2, at the top of the stack
         * op1 is left-associative and its precedence is less than
         * or equal to that of op2, or op1 is right-associative and its
         * precedence is less than that of op2 */
        if (sctoken->type == TOKEN_OPERATOR &&
            ((op_left_assoc(sctoken->data[0]) &&
             (op_preced(sctoken->data[0]) <= op_preced(sctoken->data[0]))) ||
            (!op_left_assoc(sctoken->data[0]) &&
             (op_preced(sctoken->data[0]) < op_preced(sctoken->data[0]))))) {
          /* encode operator */
          encode_operator(pkt, sctoken->data[0]);
          sl--;
        } else {
          break;
        }
      }
      /* Push operator onto the stack */
      stack[sl] = token;
      ++sl;
    }
  }

  /* When there are no more tokens to read:
   * While there are still operator tokens in the stack: */
  while(sl > 0) {
    sctoken = stack[sl - 1];
    if (sctoken->type == TOKEN_LPAREN || sctoken->type == TOKEN_RPAREN) {
      printf("Error: parentheses mismatched\n");
      return -1;
    }
    if (sctoken->type == TOKEN_NUMBER) {
      encode_number(pkt, token->data);
    } else if (sctoken->type == TOKEN_OPERATOR) {
      encode_operator(pkt, sctoken->data[0]);
    }
    --sl;
  }

#ifdef FORMULA_DEBUG
  dump_hex(pkt->data, pkt->len);
#endif

  return 0;
}

#ifdef FORMULA_DEBUG
static void dump_hex(void *vp, int length)
{
  char linebuf[80];
  int i;
  int linebuf_dirty = 0;
  unsigned char *p = (unsigned char *)vp;

  memset(linebuf, ' ', sizeof(linebuf));
  linebuf[70] = '\0';

  for (i=0; i < length; ++i) {
    int x = i % 16;
    int ch = (unsigned)p[i];
    char hex[20];
    if (x >= 8)
      x = x * 3 + 1;
    else
      x = x * 3;
    snprintf(hex, sizeof(hex), "%02x", ch);
    linebuf[x] = hex[0];
    linebuf[x+1] = hex[1];

    if (isprint(ch))
      linebuf[52+(i%16)] = ch;
    else
      linebuf[52+(i%16)] = '.';

    linebuf_dirty = 1;
    if (!((i+1)%16)) {
      printf("%s\n", linebuf);
      memset(linebuf, ' ', sizeof(linebuf));
      linebuf[70] = '\0';
      linebuf_dirty = 0;
    }
  }
  if (linebuf_dirty == 1)
    printf("%s\n", linebuf);
}
#endif

int process_formula(char *input, struct pkt *pkt)
{
  struct token_list tlist;
  struct token *token;

#ifdef FORMULA_DEBUG
  printf("Input: %s\n", input);
#endif
  TAILQ_INIT(&tlist);
  tokenize(input, &tlist);
#ifdef FORMULA_DEBUG
  printf("---\n");
  TAILQ_FOREACH(token, &tlist, tokens) {
    switch (token->type) {
    case TOKEN_EQUALS:
      printf("TOKEN_EQUALS\n");
      break;
    case TOKEN_LPAREN:
      printf("TOKEN_LPAREN\n");
      break;
    case TOKEN_RPAREN:
      printf("TOKEN_RPAREN\n");
      break;
    case TOKEN_OPERATOR:
      printf("TOKEN_OPERATOR: %s\n", token->data);
      break;
    case TOKEN_CELL:
      printf("TOKEN_CELL: %s\n", token->data);
      break;
    case TOKEN_CELLRANGE:
      printf("TOKEN_CELLRANGE: %s\n", token->data);
      break;
    case TOKEN_FUNCTION:
      printf("TOKEN_FUNCTION: %s\n", token->data);
      break;
    case TOKEN_NUMBER:
      printf("TOKEN_NUMBER: %s\n", token->data);
      break;
    case TOKEN_STRING:
      printf("TOKEN_STRING: %s\n", token->data);
      break;
    case TOKEN_COMMA:
      printf("TOKEN_COMMA\n");
      break;
    default:
      printf("Unknown token %d\n", token->type);
    }
  }
#endif
  parse_token_list(&tlist, pkt);
  while ((token = TAILQ_FIRST(&tlist))) {
    TAILQ_REMOVE(&tlist, token, tokens);
    free(token->data);
    free(token);
  }
  return 0;
}
