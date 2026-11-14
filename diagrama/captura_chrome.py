"""Captura .canvas de un HTML de stack-diagram con Chrome sin ventana (sin Playwright).
Uso: python3 captura_chrome.py IN.html OUT.png ANCHO ALTO  -> OUT.png (claro) y OUT-dark.png
Sustituye a render_png.py de stack-diagram cuando no hay Playwright: usa Google Chrome ya instalado."""
import pathlib, re, subprocess, sys, tempfile
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
src, out, w, h = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2]), int(sys.argv[3]), int(sys.argv[4])
html = src.read_text(encoding="utf-8")
css = ("<style>body{padding:0!important;margin:0!important}.head,.foot,.legend{display:none!important}"
       ".scroller{border:none!important;border-radius:0!important;overflow:visible!important}"
       ".canvas{border-radius:0!important}</style></head>")
tmp = pathlib.Path(tempfile.mkdtemp(dir=sys.argv[5] if len(sys.argv) > 5 else None))
for theme, dest in (("light", out), ("dark", out.with_name(out.stem + "-dark" + out.suffix))):
    page = re.sub(r"<html([^>]*)>", rf'<html\1 data-theme="{theme}">', html, count=1).replace("</head>", css, 1)
    p = tmp / f"{theme}.html"; p.write_text(page, encoding="utf-8")
    subprocess.run([CHROME, "--headless=new", f"--user-data-dir={tmp/'perfil'}", "--no-first-run",
        "--disable-gpu", "--hide-scrollbars", "--font-render-hinting=none", "--force-device-scale-factor=2",
        f"--window-size={w},{h}", "--virtual-time-budget=3000", f"--screenshot={dest.resolve()}", p.as_uri()],
        check=True, capture_output=True, timeout=90)
    print("rendered", dest)
