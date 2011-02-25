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

#include <sys/types.h>
#include <ctype.h>
#include <stdlib.h>
#include <string.h>

#include "stream.h"

#define MAX_SIZE	16384

struct pkt *
pkt_init(size_t len, int type)
{
	struct pkt *p = NULL;

	p = malloc((size_t) sizeof(struct pkt));
	if (p == NULL) return NULL;

	if (type == VARIABLE_PACKET) p->data = malloc(MAX_SIZE);
	else p->data = malloc(len);

	if (p->data == NULL) {
		free(p);
		return NULL;
	}

	p->len = 0;
	p->offset = 0;

	return p;

}

void pkt_addraw(struct pkt *p, const unsigned char *bytes, size_t len)
{
	memcpy(p->data + p->offset, bytes, len);
	p->offset += len;
	p->len += len;
}

void pkt_addzero(struct pkt *p, int num_zeros)
{
	memset(p->data + p->offset, 0, num_zeros);
	p->offset += num_zeros;
	p->len += num_zeros;
}

void pkt_addstring(struct pkt *p, const char *bytes) {
	uint32_t len;

	len = strlen(bytes);
	pkt_addraw(p, (unsigned char *) bytes, len);
}

void pkt_add8(struct pkt *p, uint8_t val)
{
	p->data[p->offset] = val;
	p->offset++;
	p->len++;
}

void pkt_add16_le(struct pkt *p, uint16_t val)
{
	p->data[p->offset++] = (val & 0xff);
	p->data[p->offset++] = (val & 0xff00) >> 8;
	p->len += 2;
}

void pkt_add16(struct pkt *p, uint16_t val)
{
	p->data[p->offset++] = (val & 0xff00) >> 8;
	p->data[p->offset++] = (val & 0xff);
	p->len += 2;
}

void pkt_add32_le(struct pkt *p, uint32_t val)
{
	p->data[p->offset++] = (val & 0xff);
	p->data[p->offset++] = (val & 0xff00) >> 8;
	p->data[p->offset++] = (val & 0xff0000) >> 16;
	p->data[p->offset++] = (val & 0xff000000) >> 24;
	p->len += 4;
}

void pkt_add32(struct pkt *p, uint32_t val)
{
	p->data[p->offset++] = (val & 0xff000000) >> 24;
	p->data[p->offset++] = (val & 0xff0000) >> 16;
	p->data[p->offset++] = (val & 0xff00) >> 8;
	p->data[p->offset++] = (val & 0xff);
	p->len += 4;
}

void pkt_free(struct pkt * p)
{
	if (p == NULL) return;

	if (p->data != NULL) {
		free(p->data);
	}
	free(p);
}

