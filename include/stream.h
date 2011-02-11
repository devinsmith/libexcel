/* Copyright (C) 2003-2005, Claudio Leite
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are
 * met:
 *
 *  1. Redistributions of source code must retain the above copyright
 *     notice, this list of conditions and the following disclaimer.
 *
 *  2. Redistributions in binary form must reproduce the above copyright
 *     notice, this list of conditions and the following disclaimer in the
 *     documentation and/or other materials provided with the distribution.
 *
 *  3. Neither the name of the BSF Software Project nor the names of its
 *     contributors may be used to endorse or promote products derived
 *     from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS
 * IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
 * THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
 * PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR
 * CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
 * EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
 * PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
 * PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
 * NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE
 */

/* This code is based on Claudio Leite's packet.c file inside the
 * imcomm module of his bsflite program. Claudio states that his code was
 * inspired by libfaim's (now defunct) bstream (binary stream)
 * implementation.
 */

#ifndef __STREAM_H__
#define __STREAM_H__

#include <sys/types.h>
#include <stdint.h>

#define VARIABLE_PACKET		1
#define FIXED_PACKET		2

struct pkt {
  unsigned char *data;
  size_t offset;
  size_t len;
};

struct pkt_rdr
{
	unsigned char *data;
	size_t len;
	int offset;
};

/* packet reader functions */
struct pkt_rdr * pkt_rdr_init(unsigned char *data, size_t len);
void pkt_advance_offset(struct pkt_rdr *r, int adv);
uint8_t pkt_get8(struct pkt_rdr *r);
uint16_t pkt_get16_le(struct pkt_rdr *r);
uint16_t pkt_get16(struct pkt_rdr *r);
uint32_t pkt_get32_le(struct pkt_rdr *r);
uint32_t pkt_get32(struct pkt_rdr *r);
int pkt_getstring(struct pkt_rdr *r, int len, char *str);
int pkt_getdata(struct pkt_rdr *r, int len, unsigned char *data);
void pkt_rdr_free(struct pkt_rdr *r);

/* packet creation functions */
struct pkt * pkt_tpkt_init(size_t len, int type);
void pkt_addraw(struct pkt *p, const unsigned char *bytes, size_t len);
void pkt_addzero(struct pkt *p, int num_zeros);
void pkt_add8(struct pkt *p, uint8_t val);
void pkt_add16(struct pkt *p, uint16_t val);
void pkt_add16_le(struct pkt *p, uint16_t val);
void pkt_add32(struct pkt *p, uint32_t val);
void pkt_add32_le(struct pkt *p, uint32_t val);
void pkt_write_data_len(struct pkt *p);
void pkt_addstring(struct pkt *p, const char *bytes);
void pkt_free(struct pkt * p);

/* BER specific routines */
struct pkt * pkt_ber_init(size_t len, int type, uint16_t magic_tag);
void pkt_ber_write_data(struct pkt *p, unsigned char ber_tag, int len, int data);
void pkt_ber_write_data_len(struct pkt *p);

/* Generic routines */
struct pkt * pkt_generic_init(size_t len, int type, uint16_t magic);
void pkt_generic_write_data_len(struct pkt *p);

/* Raw routines */
struct pkt * pkt_init(size_t len, int type);

#endif /* __STREAM_H__ */

