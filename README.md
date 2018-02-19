# Univeral Login

This is a prototype that connects open id (oauth2 + identity) that authenticates by pretending to be an email client and logging in via Microsoft Exchange and/or IMAP(S).  This allows almost anyone to login via their existing email address and email password. 

## Setup

The following environment varialbes must be setup:

* `KEYSTORE` - This is a keystore of certificates generated with `node ./tools/keygen.js` ran on the command line.
* `URL` - The url for this identity provider, it must be explicitly set (and not inferred from Host header) to prevent proxying attacks.
* `SECURE_KEY` - This is used for encrypting oidc information for authentication codes and transit temporary info.
* `REDIS_URL` - The redis for storing session information.
* `SECRET` - This is used for encrypting cookies.