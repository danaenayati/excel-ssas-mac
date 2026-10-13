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

Installed from PyPI into the helper's own environment; see their own licences:
pyspnego (MIT), cryptography (Apache-2.0 / BSD), keyring (MIT).
The task pane bundles React (MIT).
