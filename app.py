"""
Tienda de ropa BÁSICA - Flask + Excel como base de datos.

Ejecutar:  python app.py   (o botón ▶ de VS Code, o tecla F5)
Abrir:     http://127.0.0.1:5000
Admin:     admin@tienda.com / admin123
Datos:     datos.xlsx (hojas: Usuarios, Productos, Ventas)
"""
import os
import re
import secrets
import threading
from datetime import datetime
from functools import wraps

from flask import (Flask, abort, flash, redirect, render_template, request,
                   send_file, session, url_for)
from openpyxl import Workbook, load_workbook
from werkzeug.security import check_password_hash, generate_password_hash

CARPETA = os.environ.get("TIENDA_DATA_DIR") or os.path.dirname(os.path.abspath(__file__))
ARCHIVO = os.path.join(CARPETA, "datos.xlsx")
CANDADO = threading.Lock()

COLUMNAS = {
    "Usuarios": ["ID", "Nombre", "Email", "Password", "Rol"],
    "Productos": ["ID", "Nombre", "Precio", "Stock", "Imagen"],
    "Ventas": ["ID", "Fecha", "Cliente", "Email", "Producto", "Cantidad", "Total"],
}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 3 * 1024 * 1024  # fotos de hasta 3 MB
CARPETA_FOTOS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "img")
# Clave para firmar las sesiones. Se genera una vez y se guarda en un archivo.
os.makedirs(CARPETA, exist_ok=True)
_clave = os.path.join(CARPETA, ".clave")
if not os.path.exists(_clave):
    with open(_clave, "w") as f:
        f.write(secrets.token_hex(32))
with open(_clave) as f:
    app.secret_key = f.read().strip()


# ---------------------------------------------------------------- Excel
def crear_excel():
    """Crea datos.xlsx la primera vez, con un admin y algunos productos de ejemplo."""
    if os.path.exists(ARCHIVO):
        # Excel creado con una versión anterior: agrega la columna "Imagen" si falta.
        try:
            wb = load_workbook(ARCHIVO)
            ws = wb["Productos"]
            if ws.cell(1, 5).value != "Imagen":
                ws.cell(1, 5).value = "Imagen"
                wb.save(ARCHIVO)
            wb.close()
        except PermissionError:
            pass  # el Excel está abierto; se reintentará en el próximo arranque
        return
    wb = Workbook()
    wb.remove(wb.active)
    for hoja, cols in COLUMNAS.items():
        wb.create_sheet(hoja).append(cols)
    wb["Usuarios"].append([1, "Administrador", "admin@tienda.com",
                           generate_password_hash("admin123"), "admin"])
    ejemplos = [("Camiseta básica", 35000, 40), ("Jean slim fit", 119000, 18),
                ("Vestido floral", 135000, 12), ("Chaqueta de jean", 149000, 10),
                ("Tenis urbanos", 169000, 15), ("Gorra clásica", 29000, 30)]
    for i, (nombre, precio, stock) in enumerate(ejemplos, start=1):
        wb["Productos"].append([i, nombre, precio, stock, None])
    wb.save(ARCHIVO)


def leer(hoja):
    """Devuelve la hoja como lista de diccionarios."""
    wb = load_workbook(ARCHIVO)
    filas = [dict(zip(COLUMNAS[hoja], f))
             for f in wb[hoja].iter_rows(min_row=2, values_only=True) if f[0] is not None]
    wb.close()
    return filas


def agregar(wb, hoja, valores):
    """Añade una fila. Los textos se guardan siempre como texto (nunca como fórmula)."""
    ws = wb[hoja]
    ws.append(valores)
    for celda in ws[ws.max_row]:
        if isinstance(celda.value, str):
            celda.data_type = "s"


def nuevo_id(hoja):
    ids = [f["ID"] for f in leer(hoja)]
    return max(ids, default=0) + 1


crear_excel()


# ---------------------------------------------------------------- Utilidades
@app.template_filter("cop")
def cop(valor):
    return "$ " + "{:,.0f}".format(valor or 0).replace(",", ".")


@app.context_processor
def datos_comunes():
    if "csrf" not in session:
        session["csrf"] = secrets.token_hex(16)
    return {"csrf": session["csrf"], "sesion": session}


@app.before_request
def revisar_csrf():
    if request.method == "POST":
        token = request.form.get("csrf", "")
        if not token or not secrets.compare_digest(token, session.get("csrf", "")):
            abort(400)


def login_requerido(f):
    @wraps(f)
    def interna(*a, **k):
        if "email" not in session:
            flash("Inicia sesión para continuar.")
            return redirect(url_for("login"))
        return f(*a, **k)
    return interna


def solo_admin(f):
    @wraps(f)
    @login_requerido
    def interna(*a, **k):
        if session.get("rol") != "admin":
            abort(403)
        return f(*a, **k)
    return interna


def guardar_foto(archivo):
    """Guarda la foto subida en static/img. Devuelve el nombre del archivo o None si no es válida."""
    if not archivo or not archivo.filename:
        return None
    inicio = archivo.read(12)
    archivo.seek(0)
    if inicio.startswith(b"\x89PNG"):
        ext = "png"
    elif inicio.startswith(b"\xff\xd8\xff"):
        ext = "jpg"
    elif inicio.startswith(b"GIF8"):
        ext = "gif"
    elif inicio[:4] == b"RIFF" and inicio[8:12] == b"WEBP":
        ext = "webp"
    else:
        return None
    os.makedirs(CARPETA_FOTOS, exist_ok=True)
    nombre = secrets.token_hex(8) + "." + ext
    archivo.save(os.path.join(CARPETA_FOTOS, nombre))
    return nombre


def borrar_foto(nombre):
    """Borra una foto vieja de static/img (solo si es un archivo local)."""
    if nombre and not str(nombre).startswith("http"):
        ruta = os.path.join(CARPETA_FOTOS, os.path.basename(str(nombre)))
        if os.path.isfile(ruta):
            os.remove(ruta)


@app.errorhandler(413)
def foto_muy_grande(e):
    return "La foto es demasiado grande (máximo 3 MB). Vuelve atrás y elige una más liviana.", 413


@app.errorhandler(PermissionError)
def excel_abierto(e):
    return "No se pudo guardar. Cierra el archivo datos.xlsx si lo tienes abierto en Excel y vuelve a intentar.", 503


# ---------------------------------------------------------------- Clientes
@app.route("/")
def inicio():
    return render_template("index.html", productos=leer("Productos"))


@app.route("/registro", methods=["GET", "POST"])
def registro():
    if request.method == "POST":
        nombre = request.form.get("nombre", "").strip()
        email = request.form.get("email", "").strip().lower()
        clave = request.form.get("password", "")
        if len(nombre) < 2 or not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email) or len(clave) < 6:
            flash("Revisa los datos: nombre, correo válido y contraseña de 6+ caracteres.")
        else:
            with CANDADO:
                if any(u["Email"] == email for u in leer("Usuarios")):
                    flash("Ese correo ya está registrado.")
                else:
                    wb = load_workbook(ARCHIVO)
                    agregar(wb, "Usuarios", [nuevo_id("Usuarios"), nombre, email,
                                             generate_password_hash(clave), "cliente"])
                    wb.save(ARCHIVO)
                    flash("Cuenta creada. Ya puedes ingresar.")
                    return redirect(url_for("login"))
    return render_template("registro.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        clave = request.form.get("password", "")
        usuario = next((u for u in leer("Usuarios") if u["Email"] == email), None)
        if usuario and check_password_hash(usuario["Password"], clave):
            session.clear()
            session.update(email=usuario["Email"], nombre=usuario["Nombre"], rol=usuario["Rol"])
            return redirect(url_for("admin") if usuario["Rol"] == "admin" else url_for("inicio"))
        flash("Correo o contraseña incorrectos.")
    return render_template("login.html")


@app.route("/salir", methods=["POST"])
def salir():
    session.clear()
    return redirect(url_for("inicio"))


@app.route("/comprar/<int:pid>", methods=["POST"])
@login_requerido
def comprar(pid):
    if session["rol"] != "cliente":
        flash("Las compras se hacen con una cuenta de cliente.")
        return redirect(url_for("inicio"))
    try:
        cantidad = int(request.form.get("cantidad", 1))
    except ValueError:
        cantidad = 0
    if cantidad < 1:
        flash("Cantidad no válida.")
        return redirect(url_for("inicio"))
    with CANDADO:
        wb = load_workbook(ARCHIVO)
        ws = wb["Productos"]
        fila = next((f for f in ws.iter_rows(min_row=2) if f[0].value == pid), None)
        if fila is None:
            abort(404)
        nombre, precio, stock = fila[1].value, fila[2].value, fila[3].value
        if cantidad > stock:
            flash(f"Solo quedan {stock} unidades de {nombre}.")
            return redirect(url_for("inicio"))
        fila[3].value = stock - cantidad
        agregar(wb, "Ventas", [nuevo_id("Ventas"), datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                               session["nombre"], session["email"], nombre, cantidad,
                               precio * cantidad])
        wb.save(ARCHIVO)
    flash(f"¡Compra realizada! {cantidad} × {nombre}.")
    return redirect(url_for("mis_compras"))


@app.route("/mis-compras")
@login_requerido
def mis_compras():
    ventas = [v for v in leer("Ventas") if v["Email"] == session["email"]]
    return render_template("mis_compras.html", ventas=ventas[::-1])


# ---------------------------------------------------------------- Administración
@app.route("/admin")
@solo_admin
def admin():
    ventas = leer("Ventas")
    total = sum(v["Total"] for v in ventas)
    return render_template("admin.html", ventas=ventas[::-1], total=total,
                           productos=leer("Productos"))


@app.route("/admin/producto", methods=["POST"])
@solo_admin
def admin_producto():
    nombre = request.form.get("nombre", "").strip()
    precio = re.sub(r"\D", "", request.form.get("precio", ""))
    stock = re.sub(r"\D", "", request.form.get("stock", ""))
    archivo = request.files.get("imagen")
    foto = guardar_foto(archivo)
    if not nombre or not precio or not stock:
        flash("Completa nombre, precio y stock.")
    elif archivo and archivo.filename and not foto:
        flash("La foto debe ser PNG, JPG, GIF o WEBP.")
    else:
        with CANDADO:
            wb = load_workbook(ARCHIVO)
            agregar(wb, "Productos", [nuevo_id("Productos"), nombre, int(precio), int(stock), foto])
            wb.save(ARCHIVO)
        flash("Producto agregado.")
    return redirect(url_for("admin"))


@app.route("/admin/producto/<int:pid>/foto", methods=["POST"])
@solo_admin
def admin_foto(pid):
    foto = guardar_foto(request.files.get("imagen"))
    if not foto:
        flash("Elige una foto PNG, JPG, GIF o WEBP.")
        return redirect(url_for("admin"))
    with CANDADO:
        wb = load_workbook(ARCHIVO)
        ws = wb["Productos"]
        fila = next((f for f in ws.iter_rows(min_row=2) if f[0].value == pid), None)
        if fila is None:
            borrar_foto(foto)
            abort(404)
        celda = ws.cell(fila[0].row, 5)
        anterior = celda.value
        celda.value = foto
        celda.data_type = "s"
        wb.save(ARCHIVO)
    borrar_foto(anterior)
    flash("Foto actualizada.")
    return redirect(url_for("admin"))


@app.route("/admin/excel")
@solo_admin
def admin_excel():
    return send_file(ARCHIVO, as_attachment=True, download_name="datos.xlsx")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
