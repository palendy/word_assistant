#!/bin/bash
# Generate self-signed SSL certificates for development/internal use
mkdir -p certs
openssl req -x509 -nodes -days 365 \
  -newkey rsa:2048 \
  -keyout certs/key.pem \
  -out certs/cert.pem \
  -subj "/CN=localhost"
echo "Certificates generated in certs/"
