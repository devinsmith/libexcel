# Makefile for libexcel.

.PHONY: all clean

SRCS = biffwriter.c worksheet.c format.c formula.c hashhelp.c olewriter.c stream.c workbook.c io_handler.c

OBJS = $(SRCS:.c=.o)

CC = gcc
AR = ar

ifeq ($(OS),Windows_NT)
	RM = del /Q /F
	CP = copy /Y
else
	RM = rm -f
	CP = cp -f
endif

INTERNAL_CFLAGS = -Wall -I../include
CPPFLAGS += -MMD -MP -MT $@
CFLAGS = -O2 -pipe

LIB = libexcel
STATIC_LIB = libexcel.a

all: $(LIB)

debug: CPPFLAGS =
debug: CFLAGS = -O0 -g
debug: $(LIB)

$(LIB): $(STATIC_LIB)

$(STATIC_LIB): $(OBJS)
	$(RM) $(STATIC_LIB)
	$(AR) cru $(STATIC_LIB) $(OBJS)
	ranlib $(STATIC_LIB)

clean:
	$(RM) $(OBJS) $(STATIC_LIB)
	$(RM) *.d

.c.o:
	$(CC) $(INTERNAL_CFLAGS) $(CFLAGS) $(CPPFLAGS) -o $@ -c $<

# Include automatically generated dependency files
-include $(wildcard *.d)
