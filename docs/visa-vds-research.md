# Signed 2D barcode on visa stickers: standards, PKI and an open source implementation path

Date: 2026-11-01. Scope: the digitally signed 2D barcode printed on EU uniform format visa stickers, the standards it follows, the public key infrastructure behind it, and a pragmatic way to issue and verify such barcodes on Windows with open source tools.

Research note. This session could not reach icao.int, bsi.bund.de, publications.jrc.ec.europa.eu or secunet.com. Those hosts are outside the session's allowed network domains. Statements about the official texts therefore rest on search snippets, on the vdstools and vds-jvm reference implementations, and on prior knowledge of Doc 9303-13. Check the byte layouts against the Doc 9303 Part 13 PDF before you ship anything.

## 1. What the barcode is

The barcode is a Visible Digital Seal (VDS). A VDS is a small binary data structure that carries a subset of the printed visa data plus an ECDSA signature. The signature is made by a Visa Signer key that belongs to the issuing country's Country Signing Certification Authority (CSCA), the same root that signs ePassport chips. Anyone with the signer certificate can verify the seal offline with a camera or scanner.

The seal does not carry the signer certificate. Space on the sticker is too small. A verifier therefore needs the issuing country's Visa Signer certificates in advance. This single fact drives most of the PKI design below.

## 2. The governing documents

| Document | Role | Status |
|---|---|---|
| ICAO Doc 9303 Part 13, Visible Digital Seals, 8th edition 2021, plus Corrigendum 1 (replaces visa appendix page App A-1) | Normative encoding of the seal. Appendix A is the visa use case. Appendix C has C40 examples. | Current. Free download from icao.int. |
| ICAO Doc 9303 Part 12, Public Key Infrastructure for MRTDs, 8th edition 2021 | CSCA, signer certificate profiles, CRLs, Master Lists, ICAO PKD. VDS reuses this PKI. | Current. |
| ICAO Doc 9303 Part 7, Machine Readable Visas | Defines MRV-A and MRV-B, whose MRZ goes into the seal. | Current. |
| ICAO TR "Visible Digital Seals for Non-Electronic Documents" v1.7 | Predecessor technical report. Part 13 absorbed it. | Superseded by Part 13. |
| BSI TR-03137 Part 1 v2.5 (Dec 2021), Optically Verifiable Cryptographic Protection of non-electronic Documents (Digital Seal) | German origin of the format. Defines the German profiles (arrival attestation, residence permit, address stickers, social insurance card). BSI states the encoding is now integrated into the ICAO seal specification. | Current. |
| BSI TR-03137 Part 2 v1.0 | JAB Code, a colour barcode for larger payloads such as biometrics. Not used on visas. | Current, niche. |
| BSI TR-03171 v0.8/v0.9 | Digital seals for administrative documents. Extended VDS header, UUID profiles. German only. | Draft. Not relevant to visas. |
| ICAO TR "ICAO Datastructure for Barcode" (IDB) | Successor format. Base32 text, optional embedded signer certificate, facial image. Basis for the future EU digital visa. | Draft. |
| ICAO TR "Visible Digital Seal for Non-Constrained Environments" (VDS-NC) v1.4 | JSON seal with embedded certificate. Used for COVID health proofs. Not used on visa stickers. | Current, different family. |
| ISO 22376:2023 | Generic VDS interoperability standard with French 2D-Doc roots. States it does not replace ANTS, BSI or ICAO specifications. | Not the visa standard. |

Terminology. CSCA means Country Signing Certification Authority. DS means Document Signer. VS means Visa Signer, a DS dedicated to visa seals. PKD means the ICAO Public Key Directory. CRL means Certificate Revocation List. MRZ means Machine Readable Zone.

## 3. EU legal basis and timeline

Council Regulation (EC) 1683/95 defines the uniform format for visas. Regulation (EU) 2017/1370 introduced the current sticker design and reserved a box for a common 2D barcode because Member States could not agree on it in 2017. Agreement came in February 2020. Commission Implementing Decision C(2020) 2672 of 30 April 2020 then mandated a VDS according to Doc 9303 8th edition Parts 12 and 13, printed as a DataMatrix. The seal holds all visa data except the photograph. After a six month transition all consulates had to issue sealed stickers by May 2022. The Commission later opened infringement procedures against Bulgaria, Greece, Cyprus and Lithuania for not printing the barcode.

Regulation (EU) 2023/2685 amends 1683/95 again and moves to a fully digital visa. The digital visa is a 2D barcode cryptographically signed by the issuing Member State's CSCA and it contains the facial image. The Commission sets the application date by implementing act after it adopts the technical specifications and eu-LISA validates the arrangements. The regulation also lets the Commission classify those specifications as secret. As of this writing no start date was found. The facial image requirement means the digital visa will use the IDB structure rather than the small VDS described here. If you build for the sticker today, plan for IDB tomorrow.

## 4. Seal structure in detail

The seal has three zones. The header identifies the issuer and the signer. The message zone carries the document features as DER TLV. The signature zone carries a plain ECDSA signature over the header and the message zone. All multi-byte integers are big endian.

### 4.1 Header

| Field | Size | Encoding | Value for visas |
|---|---|---|---|
| Magic constant | 1 | fixed | 0xDC |
| Version | 1 | fixed | 0x02 = header version 3, 0x03 = header version 4. The byte is one less than the version name. |
| Issuing country | 2 | C40 of 3 chars | ICAO 3-letter code padded with '<' then '<' mapped to space, e.g. Germany is "D<<" |
| Signer identifier and certificate reference | 6 (v3) or variable (v4) | C40 | Signer identifier is 4 chars: 2-letter country code plus 2 alphanumerics, e.g. DETS. Certificate reference is the serial number of the signer certificate as a hex string. In v3 it is exactly 5 chars, padded. In v4 it is 2 hex chars giving the length, then the reference itself. |
| Document issue date | 3 | decimal MMDDYYYY as a 24-bit integer | e.g. 03/27/2024 is 03272024 = 0x31EF68 |
| Signature creation date | 3 | same | |
| Document feature definition reference | 1 | | 0x5D for visa |
| Document type category | 1 | | 0x01 for visa. Together 0x5D01 selects the visa profile. |

A v3 header is 18 bytes. A v4 header depends on the serial number length. Other profiles you will meet in the field use 0x5E03 (ICAO emergency travel document), 0xFD02 (German arrival attestation), 0xFB06 (German residence permit), 0xF908, 0xF90C and 0xF80A (German address stickers) and 0xFC04 (German social insurance card).

### 4.2 C40 encoding

C40 is the DataMatrix text compaction. Three characters become two bytes. Character values are 3 for space, 4 to 13 for digits 0 to 9, 14 to 39 for A to Z. Compute v = 1600*c1 + 40*c2 + c3 + 1 and store v as two bytes. When one character is left over, write 0xFE then the ASCII code plus one. When two are left over, pad the third with value 0. Before encoding an MRZ replace every '<' with a space. Decode in reverse: v = b0*256 + b1 - 1, c1 = v div 1600, c2 = (v mod 1600) div 40, c3 = v mod 40.

### 4.3 Message zone for the visa profile

The message zone is a sequence of DER TLV elements. The tag is one byte. The length uses DER short form below 128, 0x81 plus one byte up to 255, 0x82 plus two bytes above that. The zone ends at the first element with tag 0xFF, which is the signature.

| Tag | Feature | Encoding | Size | Required |
|---|---|---|---|---|
| 0x01 | MRZ of an MRV-A visa | C40 | 48 bytes = 72 chars: line 1 (44) plus the first 28 chars of line 2. The optional data at the end of line 2 is omitted. | One of 0x01 or 0x02 |
| 0x02 | MRZ of an MRV-B visa | C40 | 44 bytes = 64 chars: line 1 (36) plus the first 28 chars of line 2. | One of 0x01 or 0x02 |
| 0x03 | Number of entries | 1 byte unsigned integer | 1 | Yes |
| 0x04 | Duration of stay | 3 bytes: days, months, years | 3 | Yes |
| 0x05 | Passport number | C40 of 9 chars | 6 | Yes |
| 0x06 | Visa type | bytes | 1 to 4 | No |
| 0x07 | Additional features | bytes | 0 to 254 | No |

The sizes come from the vdstools VdsDocumentTypes.json profile for VISA and match the Part 13 Appendix A layout. Check the exact coding of number of entries for multiple entry visas against Appendix A because the snippet sources did not show it.

### 4.4 Signature zone

The signature element has tag 0xFF, a DER length and the plain signature r || s. There is no ASN.1 SEQUENCE around r and s and no algorithm identifier. The verifier learns the algorithm from the signer certificate. The signed data is every byte from the magic constant to the last byte of the message zone.

Part 13 chapter 2.4 ties the hash to the curve size. Use SHA-224 for fields up to 224 bits, SHA-256 up to 256 bits, SHA-384 up to 384 bits, and SHA-512 above. ICAO permits NIST and Brainpool curves. The 2023 ICAO amendments make ECDSA mandatory for VDS signers. A 256-bit curve yields a 64-byte signature, which is why 256-bit keys dominate. Germany uses brainpoolP256r1 for its seals and brainpoolP224r1 for some older German profiles.

### 4.5 Barcode

The EU decision stipulates DataMatrix (ISO/IEC 16022). Part 13 also allows Aztec and QR. Encode the seal bytes with byte compaction, not as text, because the payload is binary. A visa seal with an MRV-A MRZ and a 256-bit signature is roughly 160 to 180 bytes, which lands in a 48x48 or 52x52 square DataMatrix. Print with at least three printer dots per module. A VDS DataMatrix uses no ECI, no FNC1 and no Macro codewords.

## 5. The PKI behind the seal

### 5.1 Trust chain

The chain has two layers, the same as for ePassports. The CSCA is a self-signed root operated offline by the state. The CSCA issues Visa Signer certificates. Each Visa Signer key signs seals. The seal header names the signer by country, signer identifier and certificate serial number. The verifier resolves that triple to a certificate it already holds, checks the certificate against the CSCA and the CRL, and then verifies the seal signature.

### 5.2 Certificate profiles

Doc 9303 Part 12 defines the CSCA and DS profiles. The Visa Signer is a DS with these properties. The certificate is X.509 v3. The subject contains at least C and CN, and the CN should carry the signer identifier so that a human can map header to certificate. The serial number should be short because the v3 header gives you only 5 hex characters. Key usage is digitalSignature only. Include authorityKeyIdentifier, subjectKeyIdentifier and, where you publish one, a cRLDistributionPoints URL. Part 12 defines the ICAO documentTypeList extension (id-icao-mrtd-security-extensions-documentTypeList, OID 2.23.136.1.1.6.2). Put document type "V" there so a verifier can refuse a passport DS that is misused for visas. Confirm the exact extension content for VS certificates in Part 13 Appendix A and Part 12 before you freeze the profile.

Lifetimes follow the Part 12 Appendix A model. The CSCA private key is used for three to five years and the CSCA certificate stays valid for that period plus the longest document lifetime it must cover, so 13 to 15 years is typical. A signer private key is used for a short period, three months is the common practice, and the signer certificate stays valid for the signing period plus the maximum visa validity, which is five years for multi-entry visas, plus a safety margin. Visa seals therefore need signer certificates valid for roughly five and a half years. Issue a new signer certificate per key period rather than reusing one key.

### 5.3 Revocation and distribution

The CSCA publishes a CRL. Part 12 expects a fresh CRL at least every 90 days and immediately after a revocation. Verifiers must be able to tolerate stale CRLs because they often work offline.

Distribution of CSCA certificates is bilateral by diplomatic channel as the primary route, and via Master Lists as the secondary route. A Master List is a CMS signed list of CSCA certificates. ICAO publishes its own Master List, signed by a UN CSCA, from the CSCA certificates of PKD members. Germany, Sweden, Austria and others publish national Master Lists and their CSCA certificates on public web pages.

Distribution of Visa Signer certificates is the hard part. They are not in the seal, so they must reach border posts and consulates separately. Countries publish them alongside the CSCA. BSI's CSCA page has a "Visa Signer Certificate" section and Sweden's police site publishes VDS signer certificates. PKD members upload DS certificates to the PKD, which then serves them to other members. For EU Schengen visas the Member States exchange certificates among themselves. If you operate a verification system, build an import pipeline that accepts DER certificates, Master Lists and CRLs from several of these sources and stores them keyed by country, signer identifier and serial number.

### 5.4 Verification procedure

1. Read the DataMatrix and obtain the raw payload bytes, not the decoded string.
2. Check the magic byte 0xDC and the version byte 0x02 or 0x03.
3. Parse the header. Reject document reference values other than 0x5D01 when you expect a visa.
4. Parse the message zone up to tag 0xFF and record the signed byte range.
5. Look up the Visa Signer certificate by country, signer identifier and serial number in your trust store.
6. Validate that certificate: signature by a trusted CSCA, validity period, key usage, document type list, and absence from the CRL.
7. Check that the seal's signature creation date lies inside the signer's private key usage period and that the issue date is plausible.
8. Pick the hash from the key size and verify the plain ECDSA signature over the signed byte range.
9. Decode the features. Compare the MRZ from the seal with the MRZ printed on the sticker and read by OCR, and compare the passport number with the passport the visa sits in.
10. Treat any failure as "not verified" and display the reason. Treat a valid signature as proof of origin and integrity, not as proof that the holder is the person in front of you.

## 6. Open source building blocks

| Component | Project | Licence | Fit |
|---|---|---|---|
| Seal encode, sign, parse, verify | tsenger/vdstools, Kotlin multiplatform for JVM and iOS, Maven Central de.tsenger:vdstools 0.21.1 | Apache-2.0 | Most complete. Implements Doc 9303-13, TR-03137, TR-03171 and IDB. Builder API for signing, Verifier for plain ECDSA. Expects the public key as DER SubjectPublicKeyInfo and leaves X.509 handling to you. Companion web tool Sealgen for test seals and the Sealva app for Android and iOS to verify them. |
| Seal parse and verify, Master List reading | kurzdigital/vds-jvm, Kotlin for JVM and Android | MIT | Parses VDS and VDS-NC, verifies against a certificate list, reads CSCA Master Lists. Archived read-only in March 2026, so treat as reference code. |
| Seal parse and verify in JavaScript | li0ard/vds, TypeScript on npm and JSR | MIT | Runs in browsers, Node, Deno and Bun. Supports many curves through noble-curves including brainpoolP256r1. Useful for a browser based verifier. |
| Issuance server with HSM backends | ex_icao_vds, Elixir | Apache-2.0 | Has PKCS#11 and Vault signers and trust store adapters. Its wire format is not Doc 9303-13: it writes a 0xDC 0x04 header, a different signer reference layout and a DER signature with algorithm byte. Do not use it for interoperable visa seals. |
| Hosted verifier app | stelauconseil/vds-verify, Expo | Apache-2.0 | Sends codes to a cloud API. Not suitable for offline or sovereign use. |
| DataMatrix read with binary payload | zxing-cpp via the ZXingCpp NuGet package 0.5.1 for .NET Standard 2.0 and later, or ZXing.Net | Apache-2.0 | zxing-cpp exposes the payload bytes directly and handles binary and ECI correctly. ZXing.Net works if you concatenate ResultMetadataType.BYTE_SEGMENTS; its Result.RawBytes are codewords, not payload. |
| DataMatrix write | ZXing.Net DataMatrix writer in byte mode, or Libdmtx.Net | Apache-2.0 / LGPL | Pass the bytes as an ISO-8859-1 string or use a byte array overload so no re-encoding happens. |
| Cryptography in .NET | BouncyCastle.Cryptography | MIT | Supports Brainpool curves and the "SHA256withPLAIN-ECDSA" signer that produces and verifies r || s. Windows CNG knows Brainpool since Windows 10, but the .NET ECDsa wrappers around it are inconsistent across versions, so prefer BouncyCastle. |
| PKCS#11 from .NET | Pkcs11Interop | Apache-2.0 | Lets a .NET signing service drive SoftHSM2, Nitrokey HSM, YubiHSM or a commercial HSM. |
| Certificate authority, GUI | XCA | BSD-3 | Windows installer. Qt plus OpenSSL. Creates CSCA and signer certificates with custom extensions, manages CRLs, can use PKCS#11 tokens. Brainpool curves appear when the bundled OpenSSL lists them. Good for a small sovereign CSCA that issues a handful of signer certificates per year. |
| Certificate authority, server | EJBCA Community Edition | LGPL | Runs on Windows under Docker Desktop. Supports brainpoolP256r1, has documented CSCA and Document Signer profiles including private key usage period and the ICAO extensions. More to operate than XCA. |
| Command line crypto | OpenSSL 3 for Windows | Apache-2.0 | Scripts for key generation, certificate requests, CRLs and Master List CMS signing. Lists Brainpool curves out of the box. |
| Software HSM for development | SoftHSM2 | BSD-2 | PKCS#11 in software. Fine for test, never for production keys. |
| Affordable hardware key store | Nitrokey HSM 2 (SmartCard-HSM) with OpenSC | Open hardware and firmware, OpenSC LGPL | Supports brainpoolP256r1 and NIST curves. Works on Windows via the OpenSC PKCS#11 DLL. Match the bitness of OpenSC and your application. It is not a certified network HSM, so it suits a pilot or a small consulate footprint. |

No ready made .NET VDS library was found. The parser is small, so write it yourself and test it against seals from Sealgen and against the Sealva app.

## 7. Recommended implementation on Windows

The design below keeps the trusted parts small and offline, and keeps everything that touches the network stateless.

### 7.1 CSCA on an offline Windows machine

Use a dedicated laptop or VM that never joins a network. Install OpenSSL 3 and XCA. Generate the CSCA key on a Nitrokey HSM 2, and keep a second initialised Nitrokey in a safe as backup. Pick one curve for the whole deployment. Choose brainpoolP256r1 if your verifiers already handle German or other Brainpool seals. Choose NIST P-256 if you need the widest toolchain support, including Windows CNG. ICAO allows both.

Issue the CSCA certificate self-signed with a 13 to 15 year validity, keyCertSign and cRLSign key usage, basicConstraints CA true and pathLen 0. Produce a CRL at least every 90 days and copy it out on removable media. Publish the CSCA certificate, the CRL and all signer certificates on an HTTPS page and hand them to partners by diplomatic channel. Join the ICAO PKD when the state can.

### 7.2 Visa Signer certificates

Every three months generate a new signer key pair on the signing HSM, create a PKCS#10 request on the signing host, carry it to the CSCA machine, and sign it there. Keep the serial number small so that it still fits the v3 header. Give the certificate a validity of about five and a half years, key usage digitalSignature only, subject CN equal to the signer identifier, and the ICAO documentTypeList extension with "V". Increment the signer identifier when you add a second signing site.

### 7.3 Signing service

Build a small .NET 8 Windows service with BouncyCastle, Pkcs11Interop and ZXing.Net. Give it one job. It receives the visa fields from the visa issuing application over a local, authenticated channel, assembles the header and message zone exactly as in section 4, signs through PKCS#11 on the HSM, appends the 0xFF TLV with r || s, and returns both the raw bytes and a DataMatrix PNG. Store an audit record with the full seal bytes, the signer serial and a timestamp, because the audit trail lets you prove later what was issued. If you prefer the JVM, run vdstools behind the same interface instead of writing the encoder yourself.

### 7.4 Verification client

This fits the WIAScanner project directly. Acquire the sticker image via WIA at 300 dpi or more. Find and decode the DataMatrix with the ZXingCpp package and take the payload bytes. Parse and verify as in section 5.4 with a local trust store directory that holds CSCA certificates, Master Lists, CRLs and signer certificates. Refresh that directory from the publication sources on a schedule. Show the decoded MRZ next to the OCR MRZ so the officer sees mismatches. Never accept a seal whose signer certificate you cannot find, even if the signature math works with some other key.

### 7.5 Minimal VB.NET parser and verifier

The following code parses a visa seal and verifies the plain ECDSA signature with BouncyCastle. It is a starting point, not a certified implementation. Test it against Sealgen output.

```vb
Imports System.IO
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.Security
Imports Org.BouncyCastle.X509

Public Class VdsSeal
    Public Property IssuingCountry As String
    Public Property SignerIdentifier As String
    Public Property CertificateReference As String
    Public Property IssuingDate As Date
    Public Property SignatureDate As Date
    Public Property DocumentRef As Integer      ' &H5D01 = visa
    Public Property Features As New Dictionary(Of Byte, Byte())
    Public Property SignedBytes As Byte()       ' header + message zone
    Public Property Signature As Byte()         ' plain r || s

    Public Shared Function Parse(raw As Byte()) As VdsSeal
        If raw(0) <> &HDC Then Throw New FormatException("Not a VDS")
        Dim s As New VdsSeal, p As Integer = 2
        s.IssuingCountry = C40Decode(raw, p, 2).Replace(" "c, "<"c) : p += 2
        Select Case raw(1)
            Case &H2  ' header version 3: 4 chars signer id + 5 chars cert ref = 6 bytes
                Dim sc = C40Decode(raw, p, 6) : p += 6
                s.SignerIdentifier = sc.Substring(0, 4)
                s.CertificateReference = sc.Substring(4).Trim()
            Case &H3  ' header version 4: 4 chars signer id + 2 hex chars length + cert ref
                Dim sl = C40Decode(raw, p, 4) : p += 4
                s.SignerIdentifier = sl.Substring(0, 4)
                Dim refLen = Convert.ToInt32(sl.Substring(4, 2), 16)
                Dim refBytes = CInt(Math.Ceiling(refLen / 3.0)) * 2
                s.CertificateReference = C40Decode(raw, p, refBytes).Substring(0, refLen) : p += refBytes
            Case Else
                Throw New FormatException("Unsupported header version")
        End Select
        s.IssuingDate = DecodeDate(raw, p) : p += 3
        s.SignatureDate = DecodeDate(raw, p) : p += 3
        s.DocumentRef = (CInt(raw(p)) << 8) Or raw(p + 1) : p += 2
        While raw(p) <> &HFF
            Dim tag = raw(p) : p += 1
            Dim len = ReadDerLength(raw, p)
            s.Features(tag) = raw.Skip(p).Take(len).ToArray() : p += len
        End While
        s.SignedBytes = raw.Take(p).ToArray()
        p += 1
        Dim sigLen = ReadDerLength(raw, p)
        s.Signature = raw.Skip(p).Take(sigLen).ToArray()
        Return s
    End Function

    Public Function VerifyWith(signerCertDer As Byte()) As Boolean
        Dim cert = New X509CertificateParser().ReadCertificate(signerCertDer)
        Dim pub = DirectCast(cert.GetPublicKey(), ECPublicKeyParameters)
        Dim bits = pub.Parameters.Curve.FieldSize
        Dim algo = If(bits <= 224, "SHA224", If(bits <= 256, "SHA256", If(bits <= 384, "SHA384", "SHA512")))
        Dim verifier = SignerUtilities.GetSigner(algo & "withPLAIN-ECDSA")
        verifier.Init(False, pub)
        verifier.BlockUpdate(SignedBytes, 0, SignedBytes.Length)
        Return verifier.VerifySignature(Signature)
    End Function

    Public Function Mrz() As String
        Dim key As Byte = If(Features.ContainsKey(1), CByte(1), CByte(2))
        Return C40Decode(Features(key), 0, Features(key).Length).Replace(" "c, "<"c)
    End Function

    Private Shared Function ReadDerLength(raw As Byte(), ByRef p As Integer) As Integer
        Dim b = raw(p) : p += 1
        If b < &H80 Then Return b
        If b = &H81 Then p += 1 : Return raw(p - 1)
        If b = &H82 Then p += 2 : Return (CInt(raw(p - 2)) << 8) Or raw(p - 1)
        Throw New FormatException("Unsupported DER length")
    End Function

    Private Shared Function DecodeDate(raw As Byte(), p As Integer) As Date
        Dim v = (CInt(raw(p)) << 16) Or (CInt(raw(p + 1)) << 8) Or raw(p + 2)   ' MMDDYYYY
        Return New Date(v Mod 10000, v \ 1000000, (v Mod 1000000) \ 10000)
    End Function

    Private Shared Function C40Decode(raw As Byte(), offset As Integer, count As Integer) As String
        Dim sb As New Text.StringBuilder
        Dim i = offset
        While i < offset + count
            If raw(i) = &HFE Then
                sb.Append(ChrW(raw(i + 1) - 1)) : i += 2 : Continue While
            End If
            Dim v = (CInt(raw(i)) << 8 Or raw(i + 1)) - 1
            For Each c In {v \ 1600, (v Mod 1600) \ 40, v Mod 40}
                If c = 3 Then sb.Append(" "c)
                If c >= 4 AndAlso c <= 13 Then sb.Append(ChrW(AscW("0"c) + c - 4))
                If c >= 14 AndAlso c <= 39 Then sb.Append(ChrW(AscW("A"c) + c - 14))
            Next
            i += 2
        End While
        Return sb.ToString()
    End Function
End Class
```

Select the certificate to pass to VerifyWith by looking up IssuingCountry, SignerIdentifier and CertificateReference in your trust store. Validate the certificate chain and CRL with the .NET X509Chain class or with BouncyCastle before you call VerifyWith. Compare CertificateReference with the certificate serial number in upper case hex without leading zeros.

### 7.6 Testing

Create test seals with Sealgen and verify them with your code and with the Sealva app. Use the BSI test CSCA certificates for chain tests. Ask a neighbouring Member State for a sample sticker and its public signer certificate. Print test stickers on the real printer and read them back at the real scanner resolution. Check the failure paths: unknown signer, revoked signer, signature date outside the key usage period, altered MRZ.

## 8. Decisions you need to make

Decide the curve once, either brainpoolP256r1 or NIST P-256. Decide the header version. Version 4 is current and allows long serial numbers. Version 3 is still common. Decide whether the signing HSM is a Nitrokey class device for a pilot or a certified HSM for production. Decide where you publish certificates and CRLs and who maintains the partner exchange. Decide whether you build the verifier for VDS only or for VDS and IDB together, because the EU digital visa will arrive as IDB.

## 10. Slovenia: from the passport CSCA to visa seals

Slovenia already runs a CSCA for ePassports. Public certificate collections list it with the distinguished name C=si, O=state-institutions, OU=e-passport, CN=CSCA-Slovenia and point to csca-si.gov.si as the publication site. The Ministry of the Interior owns the passport and identity document systems, so the CSCA sits with that ministry while visas are issued by the Ministry of Foreign and European Affairs. The foreign ministry published a procurement notice around May 2022 for hardware and software to print 2D barcodes on visa stickers with 36 months of maintenance, so a first implementation most likely exists. The steps below describe what the PKI side needs regardless of whether that system is extended or replaced.

### 10.1 What the ministry needs from the CSCA

You do not need a second CSCA. The visa seal must chain to the national CSCA, the same root that signs passport Document Signer certificates. The CSCA may keep an RSA key. Part 13 mandates ECDSA for the seal signature only, so the Visa Signer key is an EC key while the certificate that carries it can be RSA-signed by the existing CSCA.

Agree a written arrangement between the two ministries, then extend the CSCA Certificate Policy and Certification Practice Statement with a Visa Signer certificate type, its naming rule, its lifetimes and its revocation procedure. Allocate one or more signer identifiers. The identifier is the two-letter code SI followed by two alphanumeric characters, for example SIV1. Register each identifier and reserve it for one signing site or one key series.

Generate every Visa Signer key pair on an HSM at the signing service of the visa system. Export a PKCS#10 request, carry it to the CSCA, and let the CSCA issue the certificate with a short serial number. Rotate the signer key roughly every three months, so the CSCA must be able to issue four to eight Visa Signer certificates per year on a schedule rather than only at passport key ceremonies.

Publish each Visa Signer certificate at the moment its key starts signing, not later. Put it next to the CSCA certificate and CRL on the CSCA publication site, hand it to the other Schengen states through the exchange channel established for Decision C(2020) 2672, and upload it to the ICAO PKD if Slovenia participates there. Keep the CRL at the 90-day cadence Part 12 expects and revoke a Visa Signer certificate immediately on key compromise.

### 10.2 Is the certificate a normal X.509 certificate?

Yes. It is a standard X.509 v3 certificate under RFC 5280 and any CA software can produce it. It must however follow the Doc 9303 Part 12 signer certificate profile, which constrains the ordinary fields rather than adding exotic ones. The public key is an EC key on a Part 13 permitted curve, in practice brainpoolP256r1 or NIST P-256. Key usage is digitalSignature only and critical. The subject distinguished name must carry the signer identifier, so use a DN such as C=SI, O=state-institutions, OU=visa, CN=SIV1. The serial number is the value the seal will reference, so keep it short and positive. Include authorityKeyIdentifier, subjectKeyIdentifier, privateKeyUsagePeriod for the signing window and a cRLDistributionPoints URL. Part 12 defines a documentTypeList extension that names the document types a signer may sign. Put the visa code V there so a verifier can refuse a passport signer that is misused for visas. Part 12 also gives a dedicated bar code signer profile in its 8th edition. Read that section and Part 13 Appendix A together before you fix the extension set, because this research could not open the PDFs.

Validity has two layers. The private key signs for about three months. The certificate itself stays valid for that window plus the longest visa validity it covers, which is five years for multiple entry visas, plus a margin, so about five and a half years.

### 10.3 What links the signature to the certificate

The seal contains no certificate, no key identifier and no hash of the certificate. The link is made by three header fields only. The issuing country, the signer identifier and the certificate reference together name exactly one certificate. Part 13 requires that the signer identifier and the certificate reference match the subject distinguished name and the serial number of the signer certificate. The verifier therefore takes the country SVN, the identifier SIV1 and the reference, which is the serial number in upper case hexadecimal without leading zeros, and looks the certificate up in its local store. In a version 3 header the reference is exactly five hex characters, so serial numbers must stay below 0x100000. In a version 4 header the reference has a two-character length prefix and can be longer.

After the lookup the verifier validates the certificate against CSCA-Slovenia and the CRL, checks that the seal's signature date falls inside the certificate's private key usage period, and then verifies the plain ECDSA signature over the header and message zone with the public key from that certificate. The successor IDB format differs here. It references the signer by five bytes of the SHA-1 hash of the certificate and can embed the certificate in the barcode, which is the model the future EU digital visa will use.

## 11. Implementation files

The folder docs/pki holds the Visa Signer certificate profile as an OpenSSL configuration, a test CSCA for rehearsals, a script that issues a test certificate, a Python script that builds and verifies a visa seal against it, and a README with the USB HSM procedure on Windows.

## 9. Sources

ICAO Doc 9303 Part 13 consolidated: https://www.icao.int/publications/Documents/9303_p13_cons_en.pdf
ICAO Doc 9303 Part 13 Corrigendum 1: https://www.icao.int/sites/default/files/publications/DocSeries/9303_p13_8ed_corr_01_en.pdf
ICAO Doc 9303 Part 13 at the ICAO store: https://store.icao.int/en/machine-readable-travel-documents-part-13-visible-digital-seals-doc-9303-part-13
ICAO Doc 9303 Part 7 consolidated: https://www.icao.int/publications/Documents/9303_p7_cons_en.pdf
ICAO Doc 9303 Part 12 consolidated: https://www.icao.int/publications/Documents/9303_p12_cons_en.pdf
ICAO TR Visible Digital Seals for Non-Electronic Documents v1.7: https://www.icao.int/Security/FAL/TRIP/Documents/TR%20-%20Visible%20Digital%20Seals%20for%20Non-Electronic%20Documents%20V1.7.pdf
ICAO TR Visible Digital Seal NC v1.4: https://www.icao.int/sites/default/files/TRIP/Publications/Technical-Report-Visible-Digital-Seal-NC-V1.4.pdf
ICAO TR ICAO Datastructure for Barcode: https://www.icao.int/sites/default/files/TRIP/Publications/ICAO-TR-ICAO-Datastructure-for-Barcode.pdf
ICAO TRIP Symposium 2023, VDS presentation (Rajeshkumar): https://www.icao.int/Meetings/TRIP-Symposium-2023/Documents/RAJESHKUMAR.pdf
ICAO TRIP Symposium 2022, VDS presentation (van Prooijen): https://www.icao.int/Meetings/TRIP-Symposium-2022/Documents/Presentations/Day2/1.VAN%20PROOIJEN.pdf
ICAO TRIP Symposium 2022, VDS presentation (Rajeshkumar): https://www.icao.int/Meetings/TRIP-Symposium-2022/Documents/Presentations/Day1/6.RAJESHKUMAR.pdf
ICAO PKD Master List Policy: https://www.icao.int/Security/FAL/PKD/Documents/ProceduresandRegulationsfortheICAOPKD/B-RoP-125-PKD%20Master%20List%20Policy.pdf
ICAO Master List page: https://www.icao.int/icao-pkd/icao-master-list
ICAO checks applied to PKD uploads: https://www.icao.int/sites/default/files/2025-06/Checks_Against_Doc9303_Applied_to_PKD_uploads.pdf
BSI TR-03137 overview: https://www.bsi.bund.de/EN/Themen/Unternehmen-und-Organisationen/Standards-und-Zertifizierung/Technische-Richtlinien/TR-nach-Thema-sortiert/tr03137/tr-03137.html
BSI TR-03137 Part 1 PDF: https://www.bsi.bund.de/SharedDocs/Downloads/EN/BSI/Publications/TechGuidelines/TR03137/BSI-TR-03137_Part1.pdf
BSI German CSCA certificates and Visa Signer section: https://www.bsi.bund.de/EN/Themen/Oeffentliche-Verwaltung/Elektronische-Identitaeten/Public-Key-Infrastrukturen/CSCA/csca.html
BSI test CSCA certificates: https://www.bsi.bund.de/EN/Themen/Oeffentliche-Verwaltung/Elektronische-Identitaeten/Public-Key-Infrastrukturen/CSCA/Test-CSCA-Zertifikate/Test-CSCA-Zertifikate.html
Swedish CSCA and VDS certificates: https://polisen.se/fr/service/passeport-et-carte-nationale/swedish-certificates-and-crls/country-signing-certificate-authority-csca-sweden/
Austrian CSCA: https://bmi.gv.at/Downloads/csca.aspx
Regulation (EU) 2017/1370: https://www.legislation.gov.uk/eur/2017/1370
European Parliament summary of the 2017 sticker amendment: https://oeil.europarl.europa.eu/oeil/en/document-summary?id=1762582
Commission proposal COM(2022) 658 on digitalisation of the visa procedure: https://ec.europa.eu/commission/presscorner/api/files/attachment/872412/COM_2022_658_1_EN_ACT_visa.pdf
Regulation (EU) 2023/2685: https://eur-lex.europa.eu/eli/reg/2023/2685/oj
Infringement on C(2020) 2672 (Greece): https://govwatch.gr/en/finds/esfalmeni-efarmogi-tis-ektelestikis-apofasis-c-2020-2672-tis-epitropis-paravasi-infr-2023-2103/
Commission action against BG, EL, CY, LT: https://schengenvisainfo.com/news/commission-requires-bulgaria-greece-cyprus-lithuania-to-implement-2d-barcode-on-uniform-format-for-visas/
JRC report Testing interoperability of VDS for Schengen visas: https://publications.jrc.ec.europa.eu/repository/handle/JRC128183
JRC report Towards testing interoperability of ICAO VDS: https://publications.jrc.ec.europa.eu/repository/handle/JRC128795
secunet Digital Seal Signer for visa factsheet: https://www.secunet.com/fileadmin/user_upload/02_Downloads/Produkt-_und_Serviceseiten_Brosch%C3%BCren_und_Factsheets/eID_PKI_Suite/secunet_DSS_for_VISA_Factsheet_EN.pdf
AKD Croatia visa VDS: https://www.akd.hr/en/solutions/government-solutions/visible-digital-seal/digitalni-pecat
Finland legislative proposal on visa digital seals: https://um.fi/press-releases/-/asset_publisher/ued5t2wDmr1C/content/viisumin-digitaaliseen-leimaan-liittyva-lakiesitys-annettu-eduskunnalle/35732
ISO 22376:2023: https://www.boutique.afnor.org/en-gb/standard/iso-223762023/security-and-resilience-authenticity-integrity-and-trust-for-products-and-d/xs142368/349832
vdstools: https://github.com/tsenger/vdstools
vdstools on openCode: https://opencode.de/en/software/vds-tools-8564
vds-jvm: https://github.com/kurzdigital/vds-jvm
li0ard/vds: https://github.com/li0ard/vds
ex_icao_vds: https://hex.pm/packages/ex_icao_vds
vds-verify: https://github.com/stelauconseil/vds-verify
ZXingCpp NuGet: https://www.nuget.org/packages/ZXingCpp
Libdmtx.Net NuGet: https://www.nuget.org/packages/Libdmtx.Net
EJBCA interoperability and supported curves: https://docs.keyfactor.com/ejbca/9.5/interoperability-and-certifications
EJBCA Document Signer profile how-to: https://docs.keyfactor.com/ejbca/9.6/create-a-certificate-profile-for-a-document-signer
EJBCA with Nitrokey HSM: https://docs.keyfactor.com/ejbca/9.4.2/nitrokey-hsm
Nitrokey HSM 2 specifications: https://shop.nitrokey.com/shop/nitrokey-hsm-2-7
python-pkcs11 notes on SmartCard-HSM curves: https://python-pkcs11.readthedocs.io/en/latest/opensc.html
XCA: https://github.com/chris2511/xca
SoftHSMv2: https://github.com/opendnssec/SoftHSMv2
