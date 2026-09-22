import os
import sqlite3
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
app.secret_key = 'clave_secreta_universidad_inventario'

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE = os.path.join(BASE_DIR, 'database.db')

def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as db:
        db.execute('''
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                rol TEXT NOT NULL
            )
        ''')
        
        db.execute('''
            CREATE TABLE IF NOT EXISTS secciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                tipo TEXT NOT NULL
            )
        ''')
        
        db.execute('''
            CREATE TABLE IF NOT EXISTS elementos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT UNIQUE NOT NULL,
                nombre TEXT NOT NULL,
                categoria TEXT NOT NULL,
                cantidad INTEGER NOT NULL DEFAULT 1,
                estado TEXT NOT NULL,
                observaciones TEXT,
                seccion_id INTEGER NOT NULL,
                FOREIGN KEY (seccion_id) REFERENCES secciones (id)
            )
        ''')
        
        cursor = db.cursor()
        
        cursor.execute("SELECT COUNT(*) FROM secciones")
        if cursor.fetchone()[0] == 0:
            db.execute("INSERT INTO secciones (nombre, tipo) VALUES ('Laboratorio 1', 'Laboratorio')")
            db.execute("INSERT INTO secciones (nombre, tipo) VALUES ('Laboratorio 2', 'Laboratorio')")
            db.execute("INSERT INTO secciones (nombre, tipo) VALUES ('Taller 1', 'Taller')")
        
        cursor.execute("SELECT COUNT(*) FROM usuarios")
        if cursor.fetchone()[0] == 0:
            db.execute("INSERT INTO usuarios (username, password, rol) VALUES ('admin', 'admin123', 'Administrador')")
            db.execute("INSERT INTO usuarios (username, password, rol) VALUES ('auxiliar', 'aux123', 'Auxiliar')")
            db.execute("INSERT INTO usuarios (username, password, rol) VALUES ('auditor', 'aud123', 'Auditor')")
            
        db.commit()

# Decorador para proteger rutas
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'username' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# --- AUTENTICACIÓN ---

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        db = get_db()
        user = db.execute('SELECT * FROM usuarios WHERE username = ? AND password = ?', (username, password)).fetchone()
        
        if user:
            session['username'] = user['username']
            session['rol'] = user['rol']
            return redirect(url_for('dashboard'))
        else:
            error = 'Usuario o contraseña incorrectos.'
            
    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# --- DASHBOARD PRINCIPAL ---

@app.route('/')
@login_required
def index():
    return redirect(url_for('dashboard'))

@app.route('/dashboard')
@login_required
def dashboard():
    db = get_db()
    secciones = db.execute("SELECT * FROM secciones").fetchall()
    
    # Resumen cuantitativo por sección
    resumen_secciones = []
    for sec in secciones:
        total = db.execute("SELECT SUM(cantidad) FROM elementos WHERE seccion_id = ?", (sec['id'],)).fetchone()[0] or 0
        operativos = db.execute("SELECT SUM(cantidad) FROM elementos WHERE seccion_id = ? AND estado = 'Operativo'", (sec['id'],)).fetchone()[0] or 0
        mantenimiento = db.execute("SELECT SUM(cantidad) FROM elementos WHERE seccion_id = ? AND estado = 'Mantenimiento'", (sec['id'],)).fetchone()[0] or 0
        inoperativos = db.execute("SELECT SUM(cantidad) FROM elementos WHERE seccion_id = ? AND estado = 'Inoperativo'", (sec['id'],)).fetchone()[0] or 0
        
        resumen_secciones.append({
            'id': sec['id'],
            'nombre': sec['nombre'],
            'tipo': sec['tipo'],
            'total_equipos': total,
            'operativos': operativos,
            'mantenimiento': mantenimiento,
            'inoperativos': inoperativos
        })
        
    # Consultar últimas observaciones reportadas
    ultimas_observaciones = db.execute('''
        SELECT e.*, s.nombre as seccion_nombre 
        FROM elementos e 
        JOIN secciones s ON e.seccion_id = s.id 
        WHERE e.observaciones IS NOT NULL AND e.observaciones != ''
        ORDER BY e.id DESC LIMIT 5
    ''').fetchall()

    return render_template('dashboard.html', 
                           secciones=secciones, 
                           resumen_secciones=resumen_secciones, 
                           ultimas_observaciones=ultimas_observaciones)

# --- INVENTARIO DE SECCIONES ---

@app.route('/seccion/<int:seccion_id>')
@login_required
def ver_seccion(seccion_id):
    db = get_db()
    secciones = db.execute("SELECT * FROM secciones").fetchall()
    seccion_actual = db.execute("SELECT * FROM secciones WHERE id = ?", (seccion_id,)).fetchone()
    elementos = db.execute("SELECT * FROM elementos WHERE seccion_id = ?", (seccion_id,)).fetchall()
    
    return render_template('inventario.html', 
                           secciones=secciones, 
                           seccion_actual=seccion_actual, 
                           elementos=elementos)

# --- OPERACIONES CRUD ---

@app.route('/agregar', methods=['POST'])
@login_required
def agregar_elemento():
    if session.get('rol') not in ['Administrador', 'Auxiliar']:
        return "Acceso denegado", 403

    codigo = request.form['codigo']
    nombre = request.form['nombre']
    categoria = request.form['categoria']
    cantidad = request.form['cantidad']
    estado = request.form['estado']
    observaciones = request.form['observaciones']
    seccion_id = request.form['seccion_id']

    db = get_db()
    try:
        db.execute('''
            INSERT INTO elementos (codigo, nombre, categoria, cantidad, estado, observaciones, seccion_id)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (codigo, nombre, categoria, cantidad, estado, observaciones, seccion_id))
        db.commit()
    except sqlite3.IntegrityError:
        print("Error: El código ya existe.")

    return redirect(url_for('ver_seccion', seccion_id=seccion_id))

@app.route('/editar/<int:elem_id>', methods=['POST'])
@login_required
def editar_elemento(elem_id):
    if session.get('rol') not in ['Administrador', 'Auxiliar']:
        return "Acceso denegado", 403

    nombre = request.form['nombre']
    categoria = request.form['categoria']
    cantidad = request.form['cantidad']
    estado = request.form['estado']
    observaciones = request.form['observaciones']
    seccion_id = request.form['seccion_id']

    db = get_db()
    db.execute('''
        UPDATE elementos 
        SET nombre = ?, categoria = ?, cantidad = ?, estado = ?, observaciones = ?
        WHERE id = ?
    ''', (nombre, categoria, cantidad, estado, observaciones, elem_id))
    db.commit()

    return redirect(url_for('ver_seccion', seccion_id=seccion_id))

@app.route('/eliminar/<int:elem_id>/<int:seccion_id>')
@login_required
def eliminar_elemento(elem_id, seccion_id):
    if session.get('rol') != 'Administrador':
        return "Acceso denegado", 403

    db = get_db()
    db.execute('DELETE FROM elementos WHERE id = ?', (elem_id,))
    db.commit()
    return redirect(url_for('ver_seccion', seccion_id=seccion_id))

if __name__ == '__main__':
    init_db()
    app.run(debug=True)