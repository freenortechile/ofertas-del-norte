# Ofertas del Norte – web

Catálogo de ofertas de Mercado Libre (tecnología, consolas, hogar, cocina y ropa) con links de afiliado. En la columna Imagen del Excel se puede pegar la dirección de la foto de Mercado Libre.

## Cómo actualizar precios
1. Abre `datos/productos.xlsx` y cambia la hoja **Productos** (precio hoy, precio antes, descuento, stock, mostrar).
2. Guarda con el mismo nombre.
3. En GitHub: carpeta `datos` → **Add file → Upload files** → arrastra `productos.xlsx` → **Commit changes**.
4. En 1-2 minutos GitHub Actions lee el Excel y publica la web actualizada.

## Archivos
- `index.html` – la página.
- `datos/productos.xlsx` – los productos (este archivo es público: no pongas ganancias ni datos personales).
- `img/` – fotos de productos y logo.
- `scripts/excel_a_json.py` – convierte el Excel en `productos.json`.
- `.github/workflows/publicar.yml` – publica en GitHub Pages en cada cambio.

## Probar en tu PC
```
pip install openpyxl
python scripts/excel_a_json.py
python -m http.server 8000
```
Luego abre http://localhost:8000
