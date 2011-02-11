# Makefile for libexcel.

.PHONY: all clean

SRCS1 = example1.c
OBJS1 = $(SRCS1:.c=.o)

SRCS2 = example2.c
OBJS2 = $(SRCS2:.c=.o)

SRCS3 = merge1.c
OBJS3 = $(SRCS3:.c=.o)

CC = gcc
AR = ar

INTERNAL_CFLAGS = -Wall -I../include
CPPFLAGS += -MMD -MP -MT $@
CFLAGS= -O2 -pipe

EXE1 = example1
EXE2 = example2
EXE3 = merge1

OBJS = $(OBJS1) $(OBJS2) $(OBJS3)
EXES = $(EXE1) $(EXE2) $(EXE3)

all: $(EXES)

$(EXE1): $(OBJS1) ../src/libexcel.a
	$(CC) -O2 -o $(EXE1) $(OBJS1) ../src/libexcel.a

$(EXE2): $(OBJS2) ../src/libexcel.a
	$(CC) -O2 -o $(EXE2) $(OBJS2) ../src/libexcel.a

$(EXE3): $(OBJS3) ../src/libexcel.a
	$(CC) -g3 -o $(EXE3) $(OBJS3) ../src/libexcel.a

clean:
	rm -f *.o $(EXES)
	rm -f *.d

.c.o:
	$(CC) $(INTERNAL_CFLAGS) $(CFLAGS) $(CPPFLAGS) -o $@ -c $<

# Include automatically generated dependency files
-include $(wildcard *.d)
