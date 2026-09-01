# -*- coding: utf-8 -*-
"""Converte .pptx em PNGs, um por slide, para o QA visual.

    python render.py output/*.pptx --out qa

Precisa de LibreOffice. O QA visual e obrigatorio: o build acerta os numeros,
mas so o render mostra tabela encostando no rodape ou texto invadindo o bloco
de baixo -- que e o que muda quando a carteira ganha ou perde um nome.
"""
import os, sys, glob, shutil, argparse, subprocess

CANDIDATOS = [
    r'C:\Program Files\LibreOffice\program\soffice.exe',
    r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    'soffice', 'libreoffice',
]


def soffice():
    for c in CANDIDATOS:
        if os.path.exists(c) or shutil.which(c):
            return c
    raise SystemExit('ERRO: LibreOffice nao encontrado. Instale ou passe --soffice.')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('arquivos', nargs='+')
    ap.add_argument('--out', default='qa')
    ap.add_argument('--dpi', type=int, default=110)
    ap.add_argument('--soffice', default=None)
    a = ap.parse_args()

    try:
        import fitz
    except ImportError:
        raise SystemExit('ERRO: falta PyMuPDF. Rode: pip install pymupdf')

    exe = a.soffice or soffice()
    os.makedirs(a.out, exist_ok=True)
    arquivos = [f for p in a.arquivos for f in glob.glob(p)]
    if not arquivos:
        raise SystemExit('ERRO: nenhum arquivo casou com o padrao')

    subprocess.run([exe, '--headless', '--convert-to', 'pdf', '--outdir', a.out] + arquivos,
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    for f in arquivos:
        base = os.path.splitext(os.path.basename(f))[0]
        pdf = os.path.join(a.out, base + '.pdf')
        if not os.path.exists(pdf):
            print('!! conversao falhou:', base)
            continue
        prefixo = ''.join(ch for ch in base.lower() if ch.isalnum())[:14]
        doc = fitz.open(pdf)
        for i, page in enumerate(doc, 1):
            saida = os.path.join(a.out, '%s-%d.png' % (prefixo, i))
            page.get_pixmap(dpi=a.dpi).save(saida)
            print(os.path.abspath(saida))
        doc.close()


if __name__ == '__main__':
    main()
