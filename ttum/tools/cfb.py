"""Minimal writer for the Compound File Binary format ([MS-CFB]) - enough to
emit a vbaProject.bin that Excel will open.

Only what a VBA project needs: version 3 (512 byte sectors), the mini stream for
small streams, a FAT that fits in the header DIFAT, and a balanced directory tree.
"""

import struct

SECTOR = 512
MINI_SECTOR = 64
MINI_CUTOFF = 4096

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF

TYPE_STORAGE = 1
TYPE_STREAM = 2
TYPE_ROOT = 5


class Entry(object):
    def __init__(self, name, kind, data=b""):
        if len(name) > 31:
            raise ValueError("directory entry names are limited to 31 characters: %s" % name)
        self.name = name
        self.kind = kind
        self.data = data
        self.children = []
        self.dir_id = None
        self.child_id = NOSTREAM
        self.left_id = NOSTREAM
        self.right_id = NOSTREAM
        self.start = ENDOFCHAIN
        self.size = 0

    def add(self, entry):
        self.children.append(entry)
        return entry


def _sort_key(entry):
    """[MS-CFB] orders siblings by name length first, then by uppercase name."""
    return (len(entry.name), entry.name.upper())


def _build_tree(entries):
    """Balanced BST over the siblings; returns the directory id of its root."""
    if not entries:
        return NOSTREAM
    mid = len(entries) // 2
    node = entries[mid]
    node.left_id = _build_tree(entries[:mid])
    node.right_id = _build_tree(entries[mid + 1:])
    return node.dir_id


def _assign_ids(root):
    """Depth-first id assignment, then a balanced sibling tree per storage."""
    ordered = []

    def walk(entry):
        entry.dir_id = len(ordered)
        ordered.append(entry)
        for child in sorted(entry.children, key=_sort_key):
            walk(child)

    walk(root)

    def link(entry):
        kids = sorted(entry.children, key=_sort_key)
        for child in kids:
            link(child)
        entry.child_id = _build_tree(kids)

    link(root)
    return ordered


def _dir_entry_bytes(entry):
    name = entry.name.encode("utf-16-le") + b"\x00\x00"
    if len(name) > 64:
        raise ValueError("directory entry name too long: %s" % entry.name)
    out = name.ljust(64, b"\x00")
    out += struct.pack("<H", len(name))
    out += struct.pack("<BB", entry.kind, 1)                  # object type, black
    out += struct.pack("<III", entry.left_id, entry.right_id, entry.child_id)
    out += b"\x00" * 16                                        # CLSID
    out += struct.pack("<I", 0)                                # state bits
    out += b"\x00" * 16                                        # creation / modified time
    out += struct.pack("<I", entry.start)
    out += struct.pack("<Q", entry.size)
    return out


def write(root):
    """Serialise the tree rooted at `root` (a TYPE_ROOT Entry) to bytes."""
    ordered = _assign_ids(root)
    streams = [e for e in ordered if e.kind == TYPE_STREAM]

    # Small streams live in the mini stream; the rest get their own sector chains.
    mini_blob = bytearray()
    mini_fat = []
    for entry in streams:
        if not entry.data or len(entry.data) >= MINI_CUTOFF:
            continue
        first = len(mini_blob) // MINI_SECTOR
        count = -(-len(entry.data) // MINI_SECTOR)
        mini_blob += entry.data.ljust(count * MINI_SECTOR, b"\x00")
        for i in range(count):
            mini_fat.append(first + i + 1 if i < count - 1 else ENDOFCHAIN)
        entry.start = first
        entry.size = len(entry.data)

    sectors = []
    fat = []

    def alloc(data):
        if not data:
            return ENDOFCHAIN
        count = -(-len(data) // SECTOR)
        first = len(sectors)
        padded = data.ljust(count * SECTOR, b"\x00")
        for i in range(count):
            sectors.append(padded[i * SECTOR:(i + 1) * SECTOR])
            fat.append(first + i + 1 if i < count - 1 else ENDOFCHAIN)
        return first

    for entry in streams:
        if entry.data and len(entry.data) >= MINI_CUTOFF:
            entry.start = alloc(entry.data)
            entry.size = len(entry.data)
        elif not entry.data:
            entry.start, entry.size = ENDOFCHAIN, 0

    root.start = alloc(bytes(mini_blob))
    root.size = len(mini_blob)

    minifat_bytes = b"".join(struct.pack("<I", v) for v in mini_fat)
    minifat_start = alloc(minifat_bytes)
    minifat_count = -(-len(minifat_bytes) // SECTOR) if minifat_bytes else 0

    dir_bytes = b"".join(_dir_entry_bytes(e) for e in ordered)
    # Directory sectors are padded with unused entries, not zeros.
    pad = (-len(dir_bytes)) % SECTOR
    if pad:
        blank = b"\x00" * 64 + struct.pack("<H", 0) + struct.pack("<BB", 0, 1) + \
                struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM) + b"\x00" * 16 + \
                struct.pack("<I", 0) + b"\x00" * 16 + struct.pack("<I", 0) + struct.pack("<Q", 0)
        dir_bytes += blank * (pad // 128)
    dir_start = alloc(dir_bytes)

    # The FAT has to describe its own sectors, so settle the count by iteration.
    n_fat = 1
    while True:
        need = -(-(len(sectors) + n_fat) // (SECTOR // 4))
        if need <= n_fat:
            break
        n_fat = need
    if n_fat > 109:
        raise ValueError("file needs a DIFAT sector chain, which this writer does not emit")

    fat_sectors = []
    for _ in range(n_fat):
        fat_sectors.append(len(sectors))
        sectors.append(b"\x00" * SECTOR)
        fat.append(FATSECT)

    fat_table = list(fat) + [FREESECT] * (n_fat * (SECTOR // 4) - len(fat))
    fat_bytes = b"".join(struct.pack("<I", v) for v in fat_table)
    for i, index in enumerate(fat_sectors):
        sectors[index] = fat_bytes[i * SECTOR:(i + 1) * SECTOR]

    header = bytearray()
    header += b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    header += b"\x00" * 16                                     # CLSID
    header += struct.pack("<HH", 0x003E, 0x0003)               # minor / major version
    header += struct.pack("<H", 0xFFFE)                        # little endian
    header += struct.pack("<HH", 9, 6)                         # 512 / 64 byte sectors
    header += b"\x00" * 6
    header += struct.pack("<I", 0)                             # directory sector count (v3: 0)
    header += struct.pack("<I", n_fat)
    header += struct.pack("<I", dir_start)
    header += struct.pack("<I", 0)                             # transaction signature
    header += struct.pack("<I", MINI_CUTOFF)
    header += struct.pack("<I", minifat_start if minifat_bytes else ENDOFCHAIN)
    header += struct.pack("<I", minifat_count)
    header += struct.pack("<I", ENDOFCHAIN)                    # no DIFAT chain
    header += struct.pack("<I", 0)
    difat = fat_sectors + [FREESECT] * (109 - n_fat)
    header += b"".join(struct.pack("<I", v) for v in difat)
    assert len(header) == SECTOR, len(header)

    return bytes(header) + b"".join(sectors)
