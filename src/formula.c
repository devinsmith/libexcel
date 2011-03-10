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

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#include "bsdqueue.h"
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

struct xl_functions {
  char *name;
  int code;
  int argc;
  int class;
};

struct xl_functions biff5_funcs[] = {
  {"SUM", 4, -1, 0},
  {"ABS", 24, 1, 1}
};

#ifdef FORMULA_DEBUG
static void dump_hex(void *vp, int length);
#endif

int is_func(char *func)
{
  int num_funcs;
  int i;

  num_funcs = sizeof(biff5_funcs) / sizeof(biff5_funcs[0]);
  for (i = 0; i < num_funcs; i++) {
    if (strcmp(func, biff5_funcs[i].name) == 0)
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
  case '-':
    pkt_add8(pkt, 0x04); /* tSub */
    break;
  case '*':
    pkt_add8(pkt, 0x05); /* tMul */
    break;
  case '/':
    pkt_add8(pkt, 0x06); /* tDiv */
    break;
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

/* Convert a cell in A1 notation to 0 based row/column notation:
 * Examples: $A$1 -> 0, 0
 * Returns: row, col, row_rel, col_rel
 */
int parse_A1(const char *cell, int *row_out, int *col_out, int *row_rel, int *col_rel)
{
  int i;
  int mul;
  int col;
  int row;
  int col_absolute;
  int row_absolute;
  int row_start;

  col_absolute = 0;
  row_absolute = 0;
  /* Check if the first character is a $. If the $ is present, then the
   * column is an absolute position. If it is missing the column is
   * relative. */
  if (cell[0] == '$') {
    col_absolute = 1;
  }

  /* Scan through the cell string looking for a $ or [0-9] */
  row_start = strcspn(cell + col_absolute, "$0123456789") + col_absolute;
  if (row_start == strlen(cell)) {
    printf("Invalid\n");
    return -1;
  }

  if (cell[row_start] == '$') {
    row_absolute = 1;
    row_start++;
  }

  mul = 0;
  col = 0;
  for (i = row_start - 1 - row_absolute; i >= col_absolute; i--) {
    if (mul == 0)
      col += (cell[i] - 65);
    else
      col += (mul * 26 * (cell[i] - 64));
    mul++;
  }

  row = strtol(cell + row_start, NULL, 10);
  row--;

  if (row_out)
    *row_out = row;
  if (col_out)
    *col_out = col;
  if (row_rel)
    *row_rel = !row_absolute;
  if (col_rel)
    *col_rel = !col_absolute;

  return 0;
}

void encode_cell(struct pkt *pkt, const char *data, int class)
{
  int row, col, c_rel, r_rel;

  if (parse_A1(data, &row, &col, &r_rel, &c_rel) == -1)
    return;

  row |= c_rel << 14;
  row |= r_rel << 15;
  pkt_add8(pkt, 0x44);  /* RefV */
  pkt_add16_le(pkt, row);
  pkt_add8(pkt, col);
#if 0
  col |= c_rel << 14;
  col |= r_rel << 15;

  pkt_add8(pkt, 0x44);  /* RefV */
  pkt_add16_le(pkt, row);
  pkt_add16_le(pkt, col);
#endif

}

void encode_function(struct pkt *pkt, const char *data, const int argc)
{
  int num_funcs;
  int i;

  num_funcs = sizeof(biff5_funcs) / sizeof(biff5_funcs[0]);
  for (i = 0; i < num_funcs; i++) {
    if (strcmp(data, biff5_funcs[i].name) == 0)
      break;
  }
  if (i == num_funcs) {
    /* This should never happen!! */
    return;
  }

  if (biff5_funcs[i].argc >= 0) {
    pkt_add8(pkt, 0x41); /* tFuncV */
  } else {
    pkt_add8(pkt, 0x42); /* tFuncVarV */
    pkt_add8(pkt, argc);
  }
  pkt_add16_le(pkt, biff5_funcs[i].code);
}

struct func_stack {
  int argc;
  int class;
};

/* Based on code from:
 * http://en.wikipedia.org/wiki/Shunting-yard_algorithm */
int parse_token_list(struct token_list *tlist, struct pkt *pkt)
{
  struct token *token;
  struct token *stack[32];
  unsigned int sl;  /* Stack length */
  struct func_stack func_stack[32]; /* Function stack */
  unsigned int fl;  /* Function length */
  struct token *sctoken;

  sl = 0;
  fl = 0;
  func_stack[fl].argc = 0;
  func_stack[fl].class = 1;
  /* Process one token at a time */
  TAILQ_FOREACH(token, tlist, tokens) {
    /* if it's a number encode it right away */
    if (token->type == TOKEN_NUMBER) {
      encode_number(pkt, token->data);
      func_stack[fl].argc++;
    }
    else if (token->type == TOKEN_CELL) {
      encode_cell(pkt, token->data, func_stack[fl].class);
      func_stack[fl].argc++;
    }
    /* If it's a function push it onto the stack */
    else if (token->type == TOKEN_FUNCTION) {
      stack[sl] = token;
      ++sl;
      func_stack[fl].argc++;
      fl++;
      func_stack[fl].argc = 0;
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
      /* When we encounter an operator, we can subtract 1 from the argument
       * count of the current function because (almost always?) the next
       * token is an operand and shouldn't count towards the argument count. */
      func_stack[fl].argc--;
    }
    /* If the token is a left parenthesis, then push it onto the stack. */
    else if (token->type == TOKEN_LPAREN) {
      stack[sl] = token;
      ++sl;
    }
    /* If the token is a right parenthesis: */
    else if (token->type == TOKEN_RPAREN) {
      int pe = 0;
      /* Until the token at the top of the stack is a left parenthesis,
       * pop operators off the stack onto the output queue */
      while (sl > 0) {
        sctoken = stack[sl - 1];
        if (sctoken->type == TOKEN_LPAREN) {
          pe = 1;
          break;
        }
        else  {
          if (sctoken->type == TOKEN_OPERATOR) {
            encode_operator(pkt, sctoken->data[0]);
          } else {
            printf("Need to encode for unknown token!\n");
          }
          sl--;
        }
      }
      /* If the stack runs out without finding a left parenthesis, then there
       * are mismatched parentheses. */
      if (!pe) {
        printf("Error: parentheses mismatched\n");
        return -1;
      }
      /* Pop the left parenthesis from the stack, but not onto the output queue. */
      sl--;
      /* If the token at the top of the stack is a function token, pop it onto
       * the output queue. */
      if (sl > 0) {
        sctoken = stack[sl - 1];
        if (sctoken->type == TOKEN_FUNCTION) {
#ifdef FORMULA_DEBUG
          printf("Arg count for function: %d\n", func_stack[fl].argc);
#endif
          encode_function(pkt, sctoken->data, func_stack[fl].argc);
          sl--;
          fl--;
        }
      }
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
    } else {
      printf("There's still something to encode\n");
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
