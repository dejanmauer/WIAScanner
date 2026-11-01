# Visa Signer certificate profile and HSM signing procedure

This folder gives you a concrete, tested starting point for issuing ICAO Doc 9303 Part 13 visa seals in Slovenia. It assumes the Ministry of Foreign and European Affairs (MFA) runs the signing service, the Ministry of the Interior runs CSCA-Slovenia, and you hold a USB HSM such as a Nitrokey HSM 2, a SmartCard-HSM or a YubiHSM 2.

Files. `visa-signer.cnf` is the certificate profile as OpenSSL configuration. It builds the PKCS#10 request on the signer side and carries the extension set the CSCA must apply. `test-csca.cnf` is a throwaway CSCA that lets you rehearse the full path before you involve the real CSCA. `make-test-pki.sh` runs that rehearsal with software keys. `make_test_seal.py` builds, signs and verifies a visa seal against the issued certificate and prints a reference vector. Run `./make-test-pki.sh` then `python3 make_test_seal.py out/vs.key out/vs.crt` to see the whole chain work. Both scripts ran successfully in this repository with OpenSSL 3.0 and brainpoolP256r1.

## 1. The profile in words, for the CSCA operator

Hand this table to the CSCA team. Their CA product will not read an OpenSSL file, but every field maps to a profile setting in EJBCA, Nexus, Entrust or whatever they run.

| Field | Value | Why |
|---|---|---|
| Version | X.509 v3 | RFC 5280, Part 12 |
| Serial number | Positive, unique, short. Keep below 0x100000 while any verifier still expects header version 3. Even number of hex digits in OpenSSL serial files. | The seal header carries this serial as the certificate reference |
| Signature algorithm | Whatever CSCA-Slovenia uses, SHA-256 or stronger. RSA is fine. | Part 13 mandates ECDSA for the seal only, not for the certificate signature |
| Issuer | The CSCA-Slovenia DN, unchanged | Chain to the existing root |
| Validity | notBefore = start of signing period. notAfter = end of signing period + 5 years + margin, about 5.5 years in total | Visas are valid up to 5 years and must stay verifiable |
| Subject | C=SI, O=state-institutions, OU=visa, CN=\<signer identifier\> | Part 12 needs C equal to the issuer's C. Part 13 needs the signer identifier recoverable from the subject DN |
| Signer identifier | SI plus two alphanumerics, for example SIV1. One per signing site or key series. | Goes into the seal header. Must be unique inside Slovenia |
| Public key | EC, brainpoolP256r1 or NIST P-256. Named curve encoding. | 256 bit gives a 64-byte signature and SHA-256 |
| keyUsage | critical, digitalSignature only | Part 12 signer profile |
| basicConstraints | absent | End-entity certificate |
| subjectKeyIdentifier | SHA-1 of the public key | Part 12 |
| authorityKeyIdentifier | keyIdentifier of the CSCA | Part 12 |
| privateKeyUsagePeriod | notBefore and notAfter of the signing window, about three months | Lets verifiers refuse seals signed outside the window |
| cRLDistributionPoints | HTTP URL of the CSCA CRL | Part 12 recommends it |
| documentTypeList (2.23.136.1.1.6.2) | version 0, set containing "V" | Part 12 extension that limits this signer to visas |
| extendedKeyUsage | absent | Not part of the Part 12 signer profile for seals. Confirm against Part 12 8th edition bar code signer section once you have the PDF |

Decode of the issued test certificate confirms the OpenSSL file produces exactly this set, including a correctly encoded `DocumentTypeListSyntax` of `30 08 02 01 00 31 03 13 01 56`.

## 2. Preparing the USB HSM on Windows

Install OpenSC. It brings `pkcs11-tool.exe`, `sc-hsm-tool.exe` and the PKCS#11 module `C:\Program Files\OpenSC Project\OpenSC\pkcs11\opensc-pkcs11.dll`. Install OpenSSL 3 for Windows. Install XCA if you prefer a GUI for the request step. Match the bitness of OpenSC, OpenSSL and your .NET process, because a 32-bit process cannot load the 64-bit module.

Initialise the token once. For a Nitrokey HSM 2 or SmartCard-HSM create a Device Key Encryption Key share first, so you can restore keys to a replacement token.

```
sc-hsm-tool --create-dkek-share dkek-share-1.pbe
sc-hsm-tool --initialize --so-pin 3537363231383830 --pin 648219 --dkek-shares 1
sc-hsm-tool --import-dkek-share dkek-share-1.pbe
```

Generate the signer key on the token. One key per signing period. Label it so that the label equals the signer identifier plus the period, which keeps audit logs readable.

```
pkcs11-tool --module opensc-pkcs11.dll --login --pin 648219 ^
  --keypairgen --key-type EC:brainpoolP256r1 --usage-sign ^
  --id 01 --label SIV1-2026Q1
pkcs11-tool --module opensc-pkcs11.dll --read-object --type pubkey --id 01 -o SIV1-2026Q1.spki.der
```

If the token refuses brainpoolP256r1, use `EC:prime256v1` and tell the CSCA that the whole deployment uses NIST P-256. Do not mix curves between signers of the same country without telling your verifiers.

## 3. Creating the request

Set `SIGNER_ID`, `KEY_PERIOD_START`, `KEY_PERIOD_END` and `CRL_URL` at the top of `visa-signer.cnf`. Then produce the PKCS#10 request with the private key on the token. OpenSSL reaches the token through the pkcs11 provider or the older libp11 engine. With the provider:

```
set PKCS11_PROVIDER_MODULE=C:\Program Files\OpenSC Project\OpenSC\pkcs11\opensc-pkcs11.dll
openssl req -new -config visa-signer.cnf -provider pkcs11 -provider default ^
  -key "pkcs11:object=SIV1-2026Q1;type=private" -out SIV1-2026Q1.csr
```

If installing the provider on Windows is troublesome, do this step in XCA. Point XCA at the OpenSC module, create a new request, pick the token key, enter the DN from the table and add the key usage. XCA cannot write the documentTypeList extension, which is fine because the CSCA applies the full extension set itself and should ignore requested extensions.

Many CSCA products also accept a bare SubjectPublicKeyInfo instead of a PKCS#10. The request format only transports the public key and proves possession. Ask the CSCA team what they want.

## 4. What the CSCA does

The CSCA imports the request, selects the Visa Signer profile from section 1, sets validity and privateKeyUsagePeriod, assigns the serial and signs. With OpenSSL the exact call is the one in `make-test-pki.sh`:

```
openssl ca -config <csca.cnf> -extfile visa-signer.cnf -extensions vs_cert_ext ^
  -startdate 20260101000000Z -enddate 20310701000000Z -in SIV1-2026Q1.csr -out SIV1-2026Q1.crt
```

Before the key signs a single visa, the CSCA publishes the certificate on its web site next to the CSCA certificate and CRL, and the MFA sends it through the Schengen exchange channel and, if Slovenia participates, to the ICAO PKD. Load the certificate back onto the token for convenience:

```
openssl x509 -in SIV1-2026Q1.crt -outform DER -out SIV1-2026Q1.der
pkcs11-tool --module opensc-pkcs11.dll --login --pin 648219 --write-object SIV1-2026Q1.der --type cert --id 01 --label SIV1-2026Q1
```

## 5. Signing seals with the HSM

The seal signature is ECDSA over the concatenation of header and message zone, with the hash chosen by curve size, and it is stored as plain r||s with each half zero-padded to the curve byte length. PKCS#11 already returns ECDSA signatures in that plain r||s form, so no DER conversion is needed. Build the bytes exactly as `make_test_seal.py` does, then sign.

```
pkcs11-tool --module opensc-pkcs11.dll --login --pin 648219 --id 01 ^
  --sign --mechanism ECDSA-SHA256 --input-file signed.bin --output-file sig.bin
```

If the token only offers raw `ECDSA`, hash the bytes with SHA-256 yourself and sign the 32-byte digest with `--mechanism ECDSA`. Do not pass `--signature-format openssl`, which would give you DER. Append `FF 40` and the 64 signature bytes to `signed.bin` and you have the seal. Encode it as a DataMatrix in byte mode.

From .NET use Pkcs11Interop. Open a session, log in with the user PIN, find the private key by label, call `Sign` with `CKM_ECDSA_SHA256` and the signed bytes, and append the returned 64 bytes under tag 0xFF. Keep the user PIN in the Windows Credential Manager or DPAPI, not in a config file. Log every signature with seal bytes, label and timestamp.

## 6. Compatibility checklist

Header country is `SVN`, C40 encoded. Signer identifier in the header equals the certificate CN byte for byte. Certificate reference equals the serial number in upper-case hexadecimal without leading zeros. If you emit header version 3, zero-pad the reference to five characters and keep serials below 0x100000. If you emit header version 4, prefix the reference with its length as two hexadecimal digits. Document feature definition reference and type category are `5D 01`. Signature creation date is the day the HSM signs. Document issue date is the visa issue date. The MRZ in tag 1 or 2 is line 1 plus the first 28 characters of line 2, with `<` replaced by space before C40. The hash is SHA-256 for a 256-bit curve. The signature is plain r||s, 64 bytes, under tag `FF 40`. The DataMatrix carries the raw bytes in byte compaction, no ECI, no text re-encoding.

Rotate the key every three months. Generate the new key and get its certificate published before the old key's usage period ends. Keep every old certificate published for at least five and a half years after its usage period, because verifiers will meet those visas at the border until they expire.

## 7. Test plan before going live

Run `make-test-pki.sh` and `make_test_seal.py` to understand the byte layout. Repeat the same steps with the real token through `pkcs11-tool` and compare the seal with the software one; only the signature bytes may differ. Decode your DataMatrix with the Sealva app and with a second independent library such as vdstools. Give your test certificate and three printed stickers to the border police so their readers can confirm the lookup by signer identifier and serial works end to end. Then repeat once more with the certificate issued by the real CSCA-Slovenia and a test visa number range.
