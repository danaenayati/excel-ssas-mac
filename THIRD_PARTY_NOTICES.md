# Third-party notices

## ssas-xmla-tcp

The SSAS wire layer in `helper/src/cubebridge/ssas/` (`dime.py`, `sealing.py`, and the
handshake and envelope logic in `session.py` and `envelopes.py`) is derived from
**ssas-xmla-tcp** by Örjan Lundberg, https://github.com/oluies/ssas-xmla-tcp, licensed
under the Apache License, Version 2.0 (http://www.apache.org/licenses/LICENSE-2.0).

That project established, against a live SQL Server 2022 Analysis Services instance,
the details Microsoft's [MS-SSAS] specification leaves out: the namespace of
`Authenticate`, the sealed-frame layout used after the handshake, and the negotiation
bits a server accepts.

Changes made here: rewritten message reader (linear time on large responses),
multidimensional (MDX cellset) execution and parsing, request properties, cancellation,
instance-name resolution, and error categories for a desktop user interface.

## Runtime dependencies

Cube Explorer itself is licensed under the GNU General Public License, version 3 or
later (see `LICENSE`). Every library below is under a licence that permits that.

Installed by `install.sh` from PyPI into the helper's private environment (not part of
this repository), each under its own licence:

| Library | Licence |
|---|---|
| pyspnego | MIT |
| gssapi (python-gssapi; macOS only) | ISC |
| decorator (used by gssapi) | BSD-2-Clause |
| krb5 (pykrb5; macOS only) | MIT |
| truststore | MIT |
| keyring (with jaraco.classes, jaraco.context, jaraco.functools, more-itertools) | MIT |
| cryptography | Apache-2.0 or BSD-3-Clause |
| cffi | MIT-0 |
| pycparser | BSD-3-Clause |

Bundled into the Excel panel when it is built: React, React DOM and Scheduler (MIT).

Tools the installer uses, downloaded from their publishers and checked against pinned
SHA-256 checksums, and not shipped with Cube Explorer: uv (Apache-2.0 or MIT), Python
(PSF License), Node.js (MIT). Build-only packages (TypeScript, Vite, Vitest and their
dependencies) are listed in `addin/package-lock.json`.
