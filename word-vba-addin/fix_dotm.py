"""
fix_dotm.py — Patches a .dotm file so its Content_Types.xml
declares 'template.macroEnabled' instead of 'document.macroEnabled'.

Word's COM automation sometimes saves .dotm with the wrong content type,
causing "Word cannot open this document template" errors.

Usage: python fix_dotm.py SuperQAT.dotm
"""
import sys, zipfile, os, shutil

def fix_dotm(path):
    tmp = path + ".tmp"
    with zipfile.ZipFile(path, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == '[Content_Types].xml':
                text = data.decode('utf-8')
                old = 'application/vnd.ms-word.document.macroEnabled.main+xml'
                new = 'application/vnd.ms-word.template.macroEnabled.main+xml'
                if old in text:
                    text = text.replace(old, new)
                    print(f"  Fixed: document.macroEnabled -> template.macroEnabled")
                    data = text.encode('utf-8')
                else:
                    print(f"  Content type already correct or not found")
            zout.writestr(item, data)
    os.replace(tmp, path)
    print(f"  Patched: {path}")

if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "SuperQAT.dotm"
    print(f"Fixing {target}...")
    fix_dotm(target)
    print("Done!")
