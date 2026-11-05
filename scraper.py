"""Scraper de muestra para el portafolio: catálogo completo de books.toscrape.com a Excel y CSV.

books.toscrape.com es un sitio hecho para practicar scraping (datos ficticios, sin login ni bloqueos).
Recorre las 50 páginas del catálogo, entra a cada libro (1000) y arma:
  - entregable/books_catalog_AAAA-MM-DD.xlsx  (hoja de libros + resumen por categoría con gráfico)
  - entregable/books_catalog_AAAA-MM-DD.csv   (UTF-8 con BOM, para que Excel muestre bien los símbolos)
"""

import csv
import datetime as dt
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin

import requests
import xlsxwriter
from lxml import html

BASE = "https://books.toscrape.com/"
PAUSA = 0.45  # segundos entre pedidos: ~2 por segundo, para no cargar el sitio
RATING = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
CARPETA = Path(__file__).resolve().parent / "entregable"

sesion = requests.Session()
sesion.headers["User-Agent"] = "Mozilla/5.0 (portfolio sample scraper)"


def bajar(url: str) -> html.HtmlElement:
    for intento in range(4):
        try:
            r = sesion.get(url, timeout=20)
            r.raise_for_status()
            r.encoding = "utf-8"
            time.sleep(PAUSA)
            return html.fromstring(r.text)
        except requests.RequestException as e:
            espera = 2 ** intento
            print(f"  reintento {intento + 1} en {espera}s: {url} ({e})", flush=True)
            time.sleep(espera)
    raise RuntimeError(f"No se pudo bajar {url}")


def urls_de_libros() -> list[str]:
    urls, pagina = [], BASE + "catalogue/page-1.html"
    while pagina:
        doc = bajar(pagina)
        for a in doc.xpath("//article[@class='product_pod']/h3/a"):
            urls.append(urljoin(pagina, a.get("href")))
        siguiente = doc.xpath("//li[@class='next']/a/@href")
        pagina = urljoin(pagina, siguiente[0]) if siguiente else None
        print(f"Catálogo: {len(urls)} libros encontrados", flush=True)
    return urls


def libro(url: str) -> dict:
    doc = bajar(url)
    tabla = {
        fila.xpath("string(th)").strip(): fila.xpath("string(td)").strip()
        for fila in doc.xpath("//table[contains(@class,'table-striped')]//tr")
    }
    estrellas = doc.xpath("//div[contains(@class,'product_main')]/p[contains(@class,'star-rating')]/@class")[0]
    stock = re.search(r"\((\d+) available\)", tabla.get("Availability", ""))
    return {
        "Title": doc.xpath("string(//div[contains(@class,'product_main')]/h1)").strip(),
        "Category": doc.xpath("string(//ul[@class='breadcrumb']/li[3]/a)").strip(),
        "Price (GBP)": float(tabla["Price (incl. tax)"].replace("£", "")),
        "Rating": RATING[estrellas.split()[-1]],
        "Stock": int(stock.group(1)) if stock else 0,
        "UPC": tabla["UPC"],
        "URL": url,
    }


def guardar_csv(filas: list[dict], ruta: Path) -> None:
    with ruta.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(filas[0]))
        w.writeheader()
        w.writerows(filas)


def guardar_excel(filas: list[dict], ruta: Path) -> None:
    libro_xl = xlsxwriter.Workbook(str(ruta))
    negrita = libro_xl.add_format({"bold": True, "bg_color": "#14271E", "font_color": "#FFFFFF", "border": 0})
    libras = libro_xl.add_format({"num_format": "£#,##0.00"})
    entero = libro_xl.add_format({"num_format": "0"})
    decimal = libro_xl.add_format({"num_format": "0.00"})
    titulo = libro_xl.add_format({"bold": True, "font_size": 14})
    nota = libro_xl.add_format({"italic": True, "font_color": "#52655A"})

    # Hoja 1: un libro por fila, como tabla de Excel (filtros y formato incluidos)
    hoja = libro_xl.add_worksheet("Books")
    columnas = list(filas[0])
    datos = [[f[c] for c in columnas] for f in filas]
    hoja.add_table(0, 0, len(datos), len(columnas) - 1, {
        "data": datos,
        "name": "Books",
        "style": "Table Style Light 15",
        "columns": [{"header": c, "header_format": negrita} for c in columnas],
    })
    for col, ancho, fmt in [(0, 55, None), (1, 20, None), (2, 12, libras), (3, 8, entero),
                            (4, 8, entero), (5, 19, None), (6, 70, None)]:
        hoja.set_column(col, col, ancho, fmt)
    hoja.freeze_panes(1, 0)

    # Hoja 2: resumen por categoría con fórmulas vivas sobre la hoja de libros
    resumen = libro_xl.add_worksheet("Summary by category")
    categorias = sorted({f["Category"] for f in filas}, key=lambda c: -sum(f["Category"] == c for f in filas))
    resumen.write(0, 0, "Books per category", titulo)
    resumen.write(1, 0, "Formulas recalculate from the Books sheet.", nota)
    encabezados = ["Category", "Books", "Average price (GBP)", "Average rating", "Units in stock"]
    for col, h in enumerate(encabezados):
        resumen.write(3, col, h, negrita)
    for i, cat in enumerate(categorias, start=4):
        fila_xl = i + 1
        resumen.write(i, 0, cat)
        resumen.write_formula(i, 1, f'=COUNTIF(Books[Category],A{fila_xl})', entero)
        resumen.write_formula(i, 2, f'=AVERAGEIF(Books[Category],A{fila_xl},Books[Price (GBP)])', libras)
        resumen.write_formula(i, 3, f'=AVERAGEIF(Books[Category],A{fila_xl},Books[Rating])', decimal)
        resumen.write_formula(i, 4, f'=SUMIF(Books[Category],A{fila_xl},Books[Stock])', entero)
    ultima = 3 + len(categorias)
    resumen.write(ultima + 1, 0, "Total", negrita)
    resumen.write_formula(ultima + 1, 1, f"=SUM(B5:B{ultima + 1})", entero)
    resumen.write_formula(ultima + 1, 4, f"=SUM(E5:E{ultima + 1})", entero)
    resumen.set_column(0, 0, 24)
    resumen.set_column(1, 4, 18)
    resumen.freeze_panes(4, 0)

    grafico = libro_xl.add_chart({"type": "bar"})
    grafico.add_series({
        "name": "Books",
        "categories": ["Summary by category", 4, 0, 13, 0],
        "values": ["Summary by category", 4, 1, 13, 1],
        "fill": {"color": "#14271E"},
        "data_labels": {"value": True},
    })
    grafico.set_title({"name": "Top 10 categories by number of books"})
    grafico.set_legend({"none": True})
    grafico.set_y_axis({"reverse": True})
    grafico.set_size({"width": 620, "height": 420})
    resumen.insert_chart("G4", grafico)
    libro_xl.close()


def main() -> int:
    CARPETA.mkdir(exist_ok=True)
    urls = urls_de_libros()
    filas = []
    for i, url in enumerate(urls, start=1):
        filas.append(libro(url))
        if i % 50 == 0 or i == len(urls):
            print(f"Libros: {i}/{len(urls)}", flush=True)
    hoy = dt.date.today().isoformat()
    guardar_csv(filas, CARPETA / f"books_catalog_{hoy}.csv")
    guardar_excel(filas, CARPETA / f"books_catalog_{hoy}.xlsx")
    print(f"Listo: {len(filas)} libros en {CARPETA}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
