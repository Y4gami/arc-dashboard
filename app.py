import sqlite3
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, jsonify

app = Flask(__name__)
DATABASE = 'arc_dashboard.db'

def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS arcs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            matricula TEXT UNIQUE NOT NULL,
            modelo TEXT NOT NULL,
            flota TEXT NOT NULL,
            arc_type TEXT NOT NULL,
            fecha_arc TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/arcs', methods=['GET', 'POST'])
def api_arcs():
    conn = get_db()
    if request.method == 'POST':
        data = request.json
        try:
            conn.execute('''
                INSERT INTO arcs (matricula, modelo, flota, arc_type, fecha_arc)
                VALUES (?, ?, ?, ?, ?)
            ''', (data['matricula'], data['modelo'], data['flota'], 
                  data['arc_type'], data['fecha_arc']))
            conn.commit()
            return jsonify({'status': 'ok'}), 201
        except sqlite3.IntegrityError:
            return jsonify({'status': 'error', 'message': 'Matrícula ya existe'}), 400
        finally:
            conn.close()
    
    arcs = conn.execute('SELECT * FROM arcs ORDER BY fecha_arc ASC').fetchall()
    conn.close()
    return jsonify([dict(row) for row in arcs])

@app.route('/api/arcs/<int:arc_id>', methods=['PUT', 'DELETE'])
def api_arc_edit(arc_id):
    conn = get_db()
    if request.method == 'PUT':
        data = request.json
        conn.execute('''
            UPDATE arcs SET matricula=?, modelo=?, flota=?, arc_type=?, fecha_arc=?
            WHERE id=?
        ''', (data['matricula'], data['modelo'], data['flota'], 
              data['arc_type'], data['fecha_arc'], arc_id))
        conn.commit()
        conn.close()
        return jsonify({'status': 'ok'})
    
    if request.method == 'DELETE':
        conn.execute('DELETE FROM arcs WHERE id=?', (arc_id,))
        conn.commit()
        conn.close()
        return jsonify({'status': 'ok'})

@app.template_filter('days_until')
def days_until(date_str):
    try:
        fecha = datetime.strptime(date_str, '%Y-%m-%d')
        return (fecha - datetime.now()).days
    except:
        return 0

@app.template_filter('urgency_class')
def urgency_class(days):
    if days < 0:
        return 'expired'
    elif days <= 30:
        return 'critical'
    elif days <= 90:
        return 'warning'
    elif days <= 180:
        return 'moderate'
    else:
        return 'safe'

if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=5000, debug=True)