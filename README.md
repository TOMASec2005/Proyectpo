# Tienda de ropa básica (Flask + Excel)

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `app.py` | Todo el programa (rutas, login, compras, panel admin) |
| `templates/` | Las páginas HTML |
| `static/style.css` | Los colores y el diseño |
| `datos.xlsx` | La base de datos (hojas Usuarios, Productos, Ventas) |
| `.vscode/launch.json` | Hace que F5 arranque la tienda en VS Code |
| `iniciar.bat` | Arranque con doble clic (Windows), sin usar VS Code |

## Requisito: tener Python

Descárgalo en https://www.python.org/downloads/ y al instalar marca **"Add Python to PATH"**.
Luego cierra y vuelve a abrir VS Code.

## Forma 1: doble clic (la más fácil, no usa terminal de VS Code)

1. Abre la carpeta `tienda_basica` en el Explorador de Windows.
2. Doble clic en **`iniciar.bat`**. Instala lo necesario y abre la tienda en el navegador.
3. Para detenerla, cierra esa ventana negra.

## Forma 2: dentro de VS Code

1. **Archivo → Abrir carpeta** y elige `tienda_basica`.
2. Instala la extensión **Python** (Microsoft) si VS Code lo pide.
3. La primera vez hay que instalar Flask y openpyxl: haz doble clic en `iniciar.bat` una vez
   (o, si tu terminal funciona: `py -m pip install -r requirements.txt`).
4. Presiona **F5** (o el botón ▶ arriba a la derecha con `app.py` abierto).
5. Abre http://127.0.0.1:5000 en el navegador.
6. Para detener: botón ■ rojo de la barra de depuración.

## Cuentas

- **Administrador:** `admin@tienda.com` / `admin123` (cámbiala en el Excel/crea otro antes de publicar).
- **Clientes:** se registran desde "Crear cuenta".

## Qué puede hacer cada uno

- **Cliente:** registrarse, ingresar, comprar productos (se descuenta el stock) y ver "Mis compras".
- **Administrador:** ver todas las ventas con total vendido, ver productos, agregar productos
  y descargar el Excel.

## Cambiar nombres, precios y fotos de los productos

- **Fotos:** entra como administrador, baja hasta "Productos y fotos", elige la imagen
  (PNG, JPG, GIF o WEBP, máximo 3 MB) y presiona **Subir**. También puedes ponerle foto
  al crear un producto nuevo. Las fotos se guardan en la carpeta `static/img`.
- **Nombres, precios y stock:** abre `datos.xlsx`, hoja **Productos**, y edita las columnas
  Nombre, Precio y Stock (no cambies la columna ID). Guarda y cierra Excel.
- Si ya tenías un `datos.xlsx` de una versión anterior, cópialo a esta carpeta: la columna
  de fotos se agrega sola al iniciar.

## Importante

Cierra `datos.xlsx` en Excel antes de vender; si está abierto, Windows no deja guardar.
