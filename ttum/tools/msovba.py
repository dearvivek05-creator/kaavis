"""MS-OVBA compression and decompression (the algorithm used by every stream in
a vbaProject.bin except _VBA_PROJECT).

Implemented from [MS-OVBA] section 2.4.1. `decompress` is here so the compressor
can be round-trip tested without Excel.
"""

import struct

MAX_CHUNK = 4096


def _ceil_log2(value: int) -> int:
    n = 0
    while (1 << n) < value:
        n += 1
    return n


def _bit_count(position: int) -> int:
    """Number of bits a copy token spends on the offset at this chunk position."""
    bits = _ceil_log2(position)
    return bits if bits >= 4 else 4


def _compress_chunk(raw: bytes) -> bytes:
    """Token-encode up to MAX_CHUNK bytes. Returns the chunk body, header excluded."""
    out = bytearray()
    pos = 0
    while pos < len(raw):
        flags = 0
        group = bytearray()
        for bit in range(8):
            if pos >= len(raw):
                break
            bits = _bit_count(pos)
            max_len = (0xFFFF >> bits) + 3
            max_off = 1 << (16 - bits)

            best_len, best_off = 0, 0
            start = pos - max_off
            if start < 0:
                start = 0
            limit = min(max_len, len(raw) - pos)
            if limit >= 3:
                for cand in range(start, pos):
                    n = 0
                    while n < limit and raw[cand + n] == raw[pos + n]:
                        n += 1
                    if n > best_len:
                        best_len, best_off = n, pos - cand
                        if n == limit:
                            break

            if best_len >= 3:
                token = ((best_off - 1) << (16 - bits)) | (best_len - 3)
                group += struct.pack("<H", token)
                flags |= 1 << bit
                pos += best_len
            else:
                group.append(raw[pos])
                pos += 1
        out.append(flags)
        out += group
    return bytes(out)


def compress(data: bytes) -> bytes:
    """Wrap `data` in a CompressedContainer."""
    if not data:
        return b"\x01"
    out = bytearray(b"\x01")
    for offset in range(0, len(data), MAX_CHUNK):
        raw = data[offset:offset + MAX_CHUNK]
        body = _compress_chunk(raw)
        if len(body) <= MAX_CHUNK and len(body) >= 1:
            header = 0xB000 | (len(body) - 1)
        else:
            # Token encoding grew the block; store it verbatim instead.
            body = raw.ljust(MAX_CHUNK, b"\x00")
            header = 0x3000 | (MAX_CHUNK - 1)
        out += struct.pack("<H", header) + body
    return bytes(out)


def decompress(data: bytes) -> bytes:
    """Inverse of `compress`, used to verify what the compressor produced."""
    if not data or data[0] != 0x01:
        raise ValueError("missing CompressedContainer signature byte")
    out = bytearray()
    i = 1
    while i < len(data):
        header = struct.unpack_from("<H", data, i)[0]
        i += 2
        size = (header & 0x0FFF) + 1
        compressed = bool(header & 0x8000)
        if (header & 0x7000) != 0x3000:
            raise ValueError("bad CompressedChunkSignature")
        chunk = data[i:i + size]
        i += size
        if not compressed:
            out += chunk
            continue
        start = len(out)
        j = 0
        while j < len(chunk):
            flags = chunk[j]
            j += 1
            for bit in range(8):
                if j >= len(chunk):
                    break
                if not (flags >> bit) & 1:
                    out.append(chunk[j])
                    j += 1
                    continue
                token = struct.unpack_from("<H", chunk, j)[0]
                j += 2
                bits = _bit_count(len(out) - start)
                length = (token & (0xFFFF >> bits)) + 3
                offset = (token >> (16 - bits)) + 1
                src = len(out) - offset
                for k in range(length):
                    out.append(out[src + k])
    return bytes(out)
