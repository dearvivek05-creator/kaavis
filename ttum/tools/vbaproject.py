"""Builds a vbaProject.bin for an Excel workbook from plain VBA source text.

Structures follow [MS-OVBA]: the `dir` stream record grammar, the PROJECT stream
properties (including the CMG/DPB/GC obfuscation from section 2.4.3), PROJECTwm,
and one compressed stream per module.
"""

import struct

import cfb
import msovba

CODEPAGE = 1252
MBCS = "cp1252"

WORKBOOK_BASE = "{00020819-0000-0000-C000-000000000046}"
WORKSHEET_BASE = "{00020820-0000-0000-C000-000000000046}"

MODULE_STANDARD = "standard"
MODULE_DOCUMENT = "document"

REFERENCES = [
    ("VBA", r"*\G{000204EF-0000-0000-C000-000000000046}#4.2#9#"
            r"C:\PROGRA~1\COMMON~1\MICROS~1\VBA\VBA7.1\VBE7.DLL#Visual Basic For Applications"),
    ("Excel", r"*\G{00020813-0000-0000-C000-000000000046}#1.9#0#"
              r"C:\PROGRA~1\MICROS~1\Office16\EXCEL.EXE#Microsoft Excel 16.0 Object Library"),
    ("stdole", r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
               r"C:\Windows\SysWOW64\stdole2.tlb#OLE Automation"),
    ("Office", r"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.8#0#"
               r"C:\PROGRA~1\COMMON~1\MICROS~1\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library"),
]


class Module(object):
    def __init__(self, name, source, kind=MODULE_STANDARD):
        self.name = name
        self.source = source
        self.kind = kind


# ---------------------------------------------------------------- dir stream

def _rec(record_id, payload):
    return struct.pack("<HI", record_id, len(payload)) + payload


def _mbcs(text):
    return text.encode(MBCS)


def _utf16(text):
    return text.encode("utf-16-le")


def _build_dir(project_name, modules):
    out = bytearray()
    out += _rec(0x0001, struct.pack("<I", 1))               # SysKind: 32-bit Windows
    out += _rec(0x0002, struct.pack("<I", 0x409))           # Lcid
    out += _rec(0x0014, struct.pack("<I", 0x409))           # LcidInvoke
    out += _rec(0x0003, struct.pack("<H", CODEPAGE))
    out += _rec(0x0004, _mbcs(project_name))
    out += _rec(0x0005, b"") + _rec(0x0040, b"")            # docstring + unicode docstring
    out += _rec(0x0006, b"") + _rec(0x003D, b"")            # help file paths
    out += _rec(0x0007, struct.pack("<I", 0))               # help context
    out += _rec(0x0008, struct.pack("<I", 0))               # lib flags
    # PROJECTVERSION carries a fixed 4-byte Reserved field where other records
    # keep their size, so it is emitted by hand.
    out += struct.pack("<HI", 0x0009, 4) + struct.pack("<I", 1) + struct.pack("<H", 0)
    out += _rec(0x000C, b"") + _rec(0x003C, b"")            # conditional compilation constants

    for name, libid in REFERENCES:
        out += struct.pack("<HI", 0x0016, len(_mbcs(name))) + _mbcs(name)
        out += struct.pack("<HI", 0x003E, len(_utf16(name))) + _utf16(name)
        body = struct.pack("<I", len(_mbcs(libid))) + _mbcs(libid) + \
               struct.pack("<I", 0) + struct.pack("<H", 0)
        out += _rec(0x000D, body)

    out += _rec(0x000F, struct.pack("<H", len(modules)))    # PROJECTMODULES
    out += _rec(0x0013, struct.pack("<H", 0xFFFF))          # PROJECTCOOKIE

    for module in modules:
        name = module.name
        out += _rec(0x0019, _mbcs(name))                    # MODULENAME
        out += _rec(0x0047, _utf16(name))                   # MODULENAMEUNICODE
        out += _rec(0x001A, _mbcs(name))                    # MODULESTREAMNAME
        out += _rec(0x0032, _utf16(name))                   #   reserved unicode form
        out += _rec(0x001C, b"") + _rec(0x0048, b"")        # MODULEDOCSTRING
        out += _rec(0x0031, struct.pack("<I", 0))           # MODULEOFFSET: no perf cache
        out += _rec(0x001E, struct.pack("<I", 0))           # MODULEHELPCONTEXT
        out += _rec(0x002C, struct.pack("<H", 0xFFFF))      # MODULECOOKIE
        out += _rec(0x0022 if module.kind == MODULE_DOCUMENT else 0x0021, b"")
        out += _rec(0x002B, b"")                            # module terminator

    out += _rec(0x0010, b"")                                # dir terminator
    return msovba.compress(bytes(out))


# ------------------------------------------------------------ PROJECT stream

def _project_key(project_id):
    return sum(_mbcs(project_id)) & 0xFF


def _encrypt(data, project_id, seed=0x0B):
    """[MS-OVBA] 2.4.3.2 - the obfuscation used for the CMG, DPB and GC values."""
    key = _project_key(project_id)
    version_enc = seed ^ 2
    key_enc = seed ^ key
    out = bytearray([seed, version_enc, key_enc])

    enc1, enc2, plain1 = version_enc, key_enc, key

    def push(value):
        nonlocal enc1, enc2, plain1
        byte = (value ^ ((enc2 + plain1) & 0xFF)) & 0xFF
        out.append(byte)
        enc2, enc1, plain1 = enc1, byte, value

    for _ in range((seed & 6) // 2):
        push(seed)
    for byte in struct.pack("<I", len(data)):
        push(byte)
    for byte in data:
        push(byte)
    return "".join("%02X" % b for b in out)


def _decrypt(hex_text, project_id):
    """Inverse of `_encrypt`, so the values written can be checked."""
    raw = bytes(int(hex_text[i:i + 2], 16) for i in range(0, len(hex_text), 2))
    seed, version_enc, key_enc = raw[0], raw[1], raw[2]
    if seed ^ version_enc != 2:
        raise ValueError("version mismatch")
    key = seed ^ key_enc
    if key != _project_key(project_id):
        raise ValueError("project key mismatch")

    enc1, enc2, plain1 = version_enc, key_enc, key
    plain = bytearray()
    for byte in raw[3:]:
        value = (byte ^ ((enc2 + plain1) & 0xFF)) & 0xFF
        plain.append(value)
        enc2, enc1, plain1 = enc1, byte, value

    ignored = (seed & 6) // 2
    length = struct.unpack_from("<I", bytes(plain), ignored)[0]
    start = ignored + 4
    return bytes(plain[start:start + length])


def _build_project(project_id, project_name, modules):
    lines = ['ID="%s"' % project_id]
    for module in modules:
        if module.kind == MODULE_DOCUMENT:
            lines.append("Document=%s/&H00000000" % module.name)
        else:
            lines.append("Module=%s" % module.name)
    lines += [
        'Name="%s"' % project_name,
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="%s"' % _encrypt(struct.pack("<I", 0), project_id),   # not protected
        'DPB="%s"' % _encrypt(b"\x00", project_id),                # no password
        'GC="%s"' % _encrypt(b"\xff", project_id),                 # project visible
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    for module in modules:
        lines.append("%s=0, 0, 0, 0, C" % module.name)
    return ("\r\n".join(lines) + "\r\n").encode(MBCS)


def _build_projectwm(modules):
    out = bytearray()
    for module in modules:
        out += _mbcs(module.name) + b"\x00"
        out += _utf16(module.name) + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)


# ------------------------------------------------------------------- module

def document_header(name, is_workbook):
    """The Attribute block Excel expects at the top of a document module."""
    base = WORKBOOK_BASE if is_workbook else WORKSHEET_BASE
    return "\r\n".join([
        'Attribute VB_Name = "%s"' % name,
        'Attribute VB_Base = "0%s"' % base,
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = False",
        "Attribute VB_PredeclaredId = True",
        "Attribute VB_Exposed = True",
        "Attribute VB_TemplateDerived = False",
        "Attribute VB_Customizable = True",
    ]) + "\r\n"


def build(project_name, modules, project_id="{2B2F1F5A-6C4E-4A9D-9E31-7C0A5D8F41B2}"):
    """Returns the bytes of a vbaProject.bin holding `modules`."""
    root = cfb.Entry("Root Entry", cfb.TYPE_ROOT)
    vba = root.add(cfb.Entry("VBA", cfb.TYPE_STORAGE))

    vba.add(cfb.Entry("_VBA_PROJECT", cfb.TYPE_STREAM,
                      b"\xcc\x61\xff\xff\x00\x00\x00"))
    vba.add(cfb.Entry("dir", cfb.TYPE_STREAM, _build_dir(project_name, modules)))
    for module in modules:
        source = module.source.replace("\r\n", "\n").replace("\n", "\r\n")
        vba.add(cfb.Entry(module.name, cfb.TYPE_STREAM,
                          msovba.compress(source.encode(MBCS))))

    root.add(cfb.Entry("PROJECT", cfb.TYPE_STREAM,
                       _build_project(project_id, project_name, modules)))
    root.add(cfb.Entry("PROJECTwm", cfb.TYPE_STREAM, _build_projectwm(modules)))

    return cfb.write(root)
