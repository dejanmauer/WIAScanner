#!/usr/bin/env bash
# Rehearse the whole issuance path with software keys:
#   test CSCA (RSA, like most legacy CSCAs) -> Visa Signer (brainpoolP256r1) -> CRL
# Run from docs/pki. Output goes to ./test-csca and ./out.
set -euo pipefail
cd "$(dirname "$0")"
CURVE="${CURVE:-brainpoolP256r1}"
SERIAL_HEX="${SERIAL_HEX:-01A3}"         # even number of hex digits; seal certificate reference becomes "1A3"
rm -rf test-csca out; mkdir -p test-csca/newcerts out
touch test-csca/index.txt
echo "$SERIAL_HEX" > test-csca/serial
echo 01 > test-csca/crlnumber

echo "== test CSCA (RSA-3072, 15 years)"
openssl genpkey -algorithm RSA -pkeyopt rsa_keygen_bits:3072 -out test-csca/csca.key 2>/dev/null
openssl req -x509 -new -config test-csca.cnf -key test-csca/csca.key -days 5475 -out test-csca/csca.crt

echo "== Visa Signer key ($CURVE). On the real system this step happens inside the HSM."
openssl genpkey -algorithm EC -pkeyopt ec_paramgen_curve:$CURVE -pkeyopt ec_param_enc:named_curve -out out/vs.key
openssl req -new -config visa-signer.cnf -key out/vs.key -out out/vs.csr
openssl req -in out/vs.csr -noout -verify -text | sed -n '1,12p'

echo "== CSCA issues the Visa Signer certificate (validity = key period + 5y + margin)"
openssl ca -batch -config test-csca.cnf -notext \
  -extfile visa-signer.cnf -extensions vs_cert_ext \
  -startdate 20260101000000Z -enddate 20310701000000Z \
  -in out/vs.csr -out out/vs.crt
openssl x509 -in out/vs.crt -outform DER -out out/vs.der
openssl x509 -in out/vs.crt -noout -text
openssl verify -CAfile test-csca/csca.crt out/vs.crt

echo "== CRL"
openssl ca -batch -config test-csca.cnf -gencrl -out test-csca/csca.crl
openssl crl -in test-csca/csca.crl -noout -text | sed -n '1,8p'

SER=$(openssl x509 -in out/vs.crt -noout -serial | cut -d= -f2 | sed 's/^0*//')
CN=$(openssl x509 -in out/vs.crt -noout -subject -nameopt sep_multiline | sed -n 's/^ *CN=//p')
echo
echo "Seal header values for this certificate:"
echo "  issuing country      : SVN"
echo "  signer identifier    : $CN"
echo "  certificate reference: $SER   (serial number, upper-case hex, no leading zeros)"
