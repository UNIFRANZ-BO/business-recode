"""Genera index.html (lo que se publica) a partir de src.html + imágenes de assets/.
Uso:  python build.py
"""
import base64, pathlib, re, shutil, subprocess, sys
raiz = pathlib.Path(__file__).parent
s = (raiz / 'src.html').read_text(encoding='utf-8')
def b64(f, mime):
    return f'data:{mime};base64,' + base64.b64encode((raiz / 'assets' / f).read_bytes()).decode()
s = s.replace('__LOGO__', b64('logo_full.png', 'image/png')).replace('__HAND__', b64('hand.jpg', 'image/jpeg'))
(raiz / 'index.html').write_text(s, encoding='utf-8')
shutil.copyfile(raiz / 'assets' / 'arte.jpg', raiz / 'og.jpg')   # vista previa de WhatsApp y «Ver la invitación original»
js = re.search(r'<script>(.*)</script>', s, re.S).group(1)
tmp = raiz / 'tests' / '.check.js'; tmp.write_text(js, encoding='utf-8')
r = subprocess.run(['node', '--check', str(tmp)], capture_output=True, text=True); tmp.unlink()
if r.returncode: print(r.stderr); sys.exit(1)
print(f'index.html generado ({len(s)//1024} KB) · sintaxis JS OK')
