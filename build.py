"""Packs addon/ into a .nvda-addon file.

An NVDA add-on is just a zip of the addon directory's contents, so this avoids
needing SCons or the add-on template for a package this small. Run with any
Python 3: ``python build.py``
"""

import os
import re
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(HERE, "addon")
MANIFEST = os.path.join(SOURCE, "manifest.ini")

#: Never ship these.
EXCLUDE_DIRS = {"__pycache__", ".git"}
EXCLUDE_SUFFIXES = (".pyc", ".pyo", ".pyd")


def readManifestValue(key):
	with open(MANIFEST, encoding="utf-8") as f:
		match = re.search(r"^%s\s*=\s*(.+)$" % re.escape(key), f.read(), re.MULTILINE)
	if not match:
		raise RuntimeError("%s missing from manifest.ini" % key)
	return match.group(1).strip().strip('"')


def build():
	name = readManifestValue("name")
	version = readManifestValue("version")
	target = os.path.join(HERE, "%s-%s.nvda-addon" % (name, version))

	written = 0
	with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zf:
		for root, dirs, files in os.walk(SOURCE):
			dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]
			for filename in files:
				if filename.endswith(EXCLUDE_SUFFIXES):
					continue
				path = os.path.join(root, filename)
				# Forward slashes, relative to addon/, or NVDA will not find anything.
				arcname = os.path.relpath(path, SOURCE).replace(os.sep, "/")
				zf.write(path, arcname)
				written += 1
	print("Wrote %s (%d files)" % (target, written))
	return target


if __name__ == "__main__":
	build()
