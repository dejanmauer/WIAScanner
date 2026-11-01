#!/usr/bin/env python3
"""Build, sign and verify a Doc 9303 Part 13 visa seal (VDS) with a software key.

Purpose: produce a reference vector that ties the certificate issued by
make-test-pki.sh to a seal, and check that header values, hash choice and the
plain r||s signature are consistent. On the real system the sign() call is
replaced by a PKCS#11 CKM_ECDSA(_SHA256) call into the HSM, which already
returns r||s.

Usage: python3 make_test_seal.py out/vs.key out/vs.crt [header_version 3|4]
"""
import sys, datetime
from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.asymmetric.utils import (
    decode_dss_signature, encode_dss_signature)

# ---------------------------------------------------------------- C40 ------
def c40_value(ch):
    if ch == ' ': return 3
    if '0' <= ch <= '9': return ord(ch) - ord('0') + 4
    if 'A' <= ch <= 'Z': return ord(ch) - ord('A') + 14
    raise ValueError(f"not C40 basic set: {ch!r}")

def c40_encode(s):
    s = s.upper().replace('<', ' ')
    out = bytearray()
    for i in range(0, len(s), 3):
        chunk = s[i:i + 3]
        if len(chunk) == 1:                      # single trailing char: FE + ascii+1
            out += bytes([0xFE, ord(chunk) + 1])
            continue
        vals = [c40_value(c) for c in chunk] + [0] * (3 - len(chunk))   # pad with 0 (shift1)
        v = 1600 * vals[0] + 40 * vals[1] + vals[2] + 1
        out += v.to_bytes(2, 'big')
    return bytes(out)

def c40_decode(b):
    out = []
    i = 0
    while i < len(b):
        if b[i] == 0xFE:
            out.append(chr(b[i + 1] - 1)); i += 2; continue
        v = int.from_bytes(b[i:i + 2], 'big') - 1
        for c in (v // 1600, (v % 1600) // 40, v % 40):
            if c == 3: out.append(' ')
            elif 4 <= c <= 13: out.append(chr(ord('0') + c - 4))
            elif 14 <= c <= 39: out.append(chr(ord('A') + c - 14))
        i += 2
    return ''.join(out)

# -------------------------------------------------------------- dates ------
def encode_date(d):                              # MMDDYYYY as 24-bit integer
    return int(f"{d.month:02d}{d.day:02d}{d.year:04d}").to_bytes(3, 'big')

def der_len(n):
    if n < 0x80: return bytes([n])
    if n < 0x100: return bytes([0x81, n])
    return bytes([0x82]) + n.to_bytes(2, 'big')

def tlv(tag, value):
    return bytes([tag]) + der_len(len(value)) + value

# ----------------------------------------------------------------- MRZ ------
def check_digit(s):
    w = (7, 3, 1); t = 0
    for i, ch in enumerate(s):
        v = 0 if ch == '<' else (int(ch) if ch.isdigit() else ord(ch) - 55)
        t += v * w[i % 3]
    return str(t % 10)

def mrva(surname, given, passport_no, nationality, dob, sex, expiry):
    l1 = ("V<SVN" + surname + "<<" + given).ljust(44, '<')[:44]
    pn = passport_no.ljust(9, '<')
    l2 = pn + check_digit(pn) + nationality + dob + check_digit(dob) + sex + expiry + check_digit(expiry)
    l2 = l2.ljust(44, '<')
    return l1, l2

# ---------------------------------------------------------------- seal ------
def build_visa_seal(signer_key, cert, header_version=4, issue=None, sig_date=None):
    country = "SVN"
    signer_id = cert.subject.get_attributes_for_oid(x509.NameOID.COMMON_NAME)[0].value
    cert_ref = format(cert.serial_number, 'X')          # upper-case hex, no leading zeros
    issue = issue or datetime.date.today()
    sig_date = sig_date or datetime.date.today()

    if header_version == 3:
        if len(cert_ref) > 5: raise ValueError("v3 header allows at most 5 hex chars of serial")
        sid = signer_id + cert_ref.rjust(5, '0')         # 9 chars -> 6 bytes
        ver = 0x02
    else:
        sid = signer_id + format(len(cert_ref), '02X') + cert_ref
        ver = 0x03
    header = (bytes([0xDC, ver]) + c40_encode(country) + c40_encode(sid)
              + encode_date(issue) + encode_date(sig_date) + bytes([0x5D, 0x01]))

    l1, l2 = mrva("NOVAK", "JANEZ", "P12345678", "SVN", "850101", "M", "300101")
    mrz72 = l1 + l2[:28]                                   # Part 13 App. A: line 1 + first 28 of line 2
    messages = (tlv(0x01, c40_encode(mrz72))               # MRZ of MRV-A, 48 bytes
                + tlv(0x03, bytes([0x01]))                 # number of entries (check App. A for MULT coding)
                + tlv(0x04, bytes([90, 0, 0]))             # duration of stay: 90 days, 0 months, 0 years
                + tlv(0x05, c40_encode("P12345678"))       # passport number, 9 chars -> 6 bytes
                + tlv(0x06, b"C"))                         # visa type (optional)
    signed = header + messages

    bits = signer_key.curve.key_size
    digest = (hashes.SHA224() if bits <= 224 else hashes.SHA256() if bits <= 256
              else hashes.SHA384() if bits <= 384 else hashes.SHA512())
    der_sig = signer_key.sign(signed, ec.ECDSA(digest))
    r, s = decode_dss_signature(der_sig)
    n = (bits + 7) // 8
    plain = r.to_bytes(n, 'big') + s.to_bytes(n, 'big')   # plain r||s, fixed width
    return signed + tlv(0xFF, plain), digest

def parse_and_verify(seal, cert):
    assert seal[0] == 0xDC
    p = 2
    country = c40_decode(seal[p:p + 2]); p += 2
    if seal[1] == 0x02:
        sc = c40_decode(seal[p:p + 6]); p += 6
        signer_id, cert_ref = sc[:4], sc[4:].lstrip('0')
    elif seal[1] == 0x03:
        sl = c40_decode(seal[p:p + 4]); p += 4
        signer_id, n = sl[:4], int(sl[4:6], 16)
        nb = -(-n // 3) * 2
        cert_ref = c40_decode(seal[p:p + nb])[:n]; p += nb
    else:
        raise ValueError("unknown header version")
    def dec_date(b):
        v = int.from_bytes(b, 'big'); return datetime.date(v % 10000, v // 1000000, (v % 1000000) // 10000)
    issue = dec_date(seal[p:p + 3]); p += 3
    sdate = dec_date(seal[p:p + 3]); p += 3
    docref = seal[p] << 8 | seal[p + 1]; p += 2
    feats = {}
    while seal[p] != 0xFF:
        tag = seal[p]; p += 1
        ln = seal[p]; p += 1
        if ln == 0x81: ln = seal[p]; p += 1
        elif ln == 0x82: ln = seal[p] << 8 | seal[p + 1]; p += 2
        feats[tag] = seal[p:p + ln]; p += ln
    signed = seal[:p]
    p += 1
    ln = seal[p]; p += 1
    sig = seal[p:p + ln]

    # certificate lookup check: header must name exactly this certificate
    cn = cert.subject.get_attributes_for_oid(x509.NameOID.COMMON_NAME)[0].value
    assert signer_id == cn, (signer_id, cn)
    assert cert_ref == format(cert.serial_number, 'X'), (cert_ref, cert.serial_number)
    assert docref == 0x5D01

    pub = cert.public_key()
    half = len(sig) // 2
    der = encode_dss_signature(int.from_bytes(sig[:half], 'big'), int.from_bytes(sig[half:], 'big'))
    bits = pub.curve.key_size
    digest = (hashes.SHA224() if bits <= 224 else hashes.SHA256() if bits <= 256
              else hashes.SHA384() if bits <= 384 else hashes.SHA512())
    pub.verify(der, signed, ec.ECDSA(digest))        # raises on failure
    return dict(country=country, signer_id=signer_id, cert_ref=cert_ref, issue=issue,
                sig_date=sdate, docref=hex(docref), mrz=c40_decode(feats[1]).replace(' ', '<'),
                entries=feats[3][0], stay=list(feats[4]), passport=c40_decode(feats[5]).replace(' ', '<'),
                visa_type=feats.get(6), sig_len=len(sig), total_len=len(seal))

if __name__ == "__main__":
    key_path, crt_path = sys.argv[1], sys.argv[2]
    hv = int(sys.argv[3]) if len(sys.argv) > 3 else 4
    key = serialization.load_pem_private_key(open(key_path, 'rb').read(), None)
    cert = x509.load_pem_x509_certificate(open(crt_path, 'rb').read())
    seal, digest = build_visa_seal(key, cert, hv,
                                   issue=datetime.date(2026, 2, 10), sig_date=datetime.date(2026, 2, 10))
    print(f"header version {hv}, hash {digest.name}, seal {len(seal)} bytes")
    print(seal.hex().upper())
    open(f"out/visa-seal-v{hv}.bin", "wb").write(seal)
    for k, v in parse_and_verify(seal, cert).items():
        print(f"  {k:10}: {v}")
    print("signature verified against", crt_path)
