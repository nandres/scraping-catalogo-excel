# Scraping de 1000 productos a Excel

Muestra de trabajo. Un scraper en Python recorre las 50 páginas de un catálogo y entra a cada uno de los 1000 productos para sacar título, categoría, precio, puntaje, stock, código y enlace.

Entrega un Excel listo para filtrar, con un resumen por categoría que se recalcula solo y un gráfico, más los mismos datos en CSV. Va a un ritmo moderado para no cargar el sitio y reintenta si una página falla.

Está hecho sobre [books.toscrape.com](https://books.toscrape.com), un sitio público pensado para practicar scraping.

```bash
pip install -r requirements.txt
python scraper.py
```

El resultado de ejemplo está en `entregable/`.

Hecho por Marcos Arellano · [marcos-arellano.vercel.app](https://marcos-arellano.vercel.app)
