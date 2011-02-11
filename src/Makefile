# Makefile for libexcel.

.PHONY: all clean

SRCS = biffwriter.c worksheet.c format.c hashhelp.c olewriter.c stream.c workbook.c

OBJS = $(SRCS:.c=.o)

CC = gcc
AR = ar

INTERNAL_CFLAGS = -Wall -I../include
CPPFLAGS += -MMD -MP -MT $@
CFLAGS= -O2 -pipe

LIB = libexcel
STATIC_LIB = libexcel.a

all: $(LIB)

$(LIB): $(STATIC_LIB)

$(STATIC_LIB): $(OBJS)
	@rm -f $(STATIC_LIB)
	$(AR) cru $(STATIC_LIB) $(OBJS)
	ranlib $(STATIC_LIB)

clean:
	rm -f $(OBJS) $(STATIC_LIB)
	rm -f *.d

.c.o:
	$(CC) $(INTERNAL_CFLAGS) $(CFLAGS) $(CPPFLAGS) -o $@ -c $<

# Include automatically generated dependency files
-include $(wildcard *.d)
