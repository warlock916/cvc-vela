from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import sqlite3, os, hashlib, secrets, functools

app = Flask(__name__, static_folder='static')
CORS(app)

DB = os.environ.get('DB_PATH', 'cvc.db')
ADMIN_PWD = os.environ.get('ADMIN_PASSWORD', 'admin123')

# ── Database ──────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as db:
        db.execute('''CREATE TABLE IF NOT EXISTS valutazioni (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            data      TEXT    NOT NULL,
            istruttore TEXT   NOT NULL,
            allievo   TEXT    NOT NULL,
            corso     TEXT    NOT NULL,
            tec_g1 INTEGER, tec_g2 INTEGER, tec_g3 INTEGER,
            tec_g4 INTEGER, tec_g5 INTEGER, tec_g6 INTEGER, tec_g7 INTEGER,
            sen_g1 INTEGER, sen_g2 INTEGER, sen_g3 INTEGER,
            sen_g4 INTEGER, sen_g5 INTEGER, sen_g6 INTEGER, sen_g7 INTEGER,
            aff_g1 INTEGER, aff_g2 INTEGER, aff_g3 INTEGER,
            aff_g4 INTEGER, aff_g5 INTEGER, aff_g6 INTEGER, aff_g7 INTEGER,
            pro_g1 INTEGER, pro_g2 INTEGER, pro_g3 INTEGER,
            pro_g4 INTEGER, pro_g5 INTEGER, pro_g6 INTEGER, pro_g7 INTEGER,
            imp_g1 INTEGER, imp_g2 INTEGER, imp_g3 INTEGER,
            imp_g4 INTEGER, imp_g5 INTEGER, imp_g6 INTEGER, imp_g7 INTEGER,
            dis_g1 INTEGER, dis_g2 INTEGER, dis_g3 INTEGER,
            dis_g4 INTEGER, dis_g5 INTEGER, dis_g6 INTEGER, dis_g7 INTEGER,
            com_g1 INTEGER, com_g2 INTEGER, com_g3 INTEGER,
            com_g4 INTEGER, com_g5 INTEGER, com_g6 INTEGER, com_g7 INTEGER,
            pts_g1 INTEGER, pts_g2 INTEGER, pts_g3 INTEGER,
            pts_g4 INTEGER, pts_g5 INTEGER, pts_g6 INTEGER, pts_g7 INTEGER,
            punteggio_finale REAL
        )''')
        db.execute('''CREATE TABLE IF NOT EXISTS sessions (
            token TEXT PRIMARY KEY,
            created_at TEXT DEFAULT (datetime('now'))
        )''')
        db.commit()

init_db()

# ── Calcolo punteggio (stessa logica del foglio Excel) ────────────────────
PESI = {
    'D1':[0.30,0.10,0.10,0.15,0.12,0.13,0.10],
    'D2':[0.30,0.10,0.10,0.15,0.12,0.13,0.10],
    'D3':[0.30,0.15,0.15,0.10,0.08,0.12,0.10],
    'D4':[0.30,0.15,0.20,0.05,0.08,0.12,0.10],
    'D5':[0.30,0.15,0.24,0.05,0.08,0.08,0.10],
    'C1':[0.35,0.15,0.10,0.10,0.10,0.10,0.10],
    'C2':[0.35,0.15,0.10,0.10,0.10,0.10,0.10],
    'C3':[0.30,0.20,0.15,0.05,0.10,0.10,0.10],
    'C4':[0.25,0.25,0.15,0.05,0.10,0.10,0.10],
    'C5':[0.25,0.25,0.15,0.05,0.10,0.10,0.10],
}

def calcola_punteggio(corso, voti):
    """voti = [tec, sen, aff, pro, imp, dis, com]"""
    p = PESI.get(corso)
    if not p or None in voti: return None
    tec, sen, aff, pro, imp, dis, com = voti
    vp = sum(v*w for v,w in zip(voti, p))
    if corso in ('D1','D2','C1','C2'):
        if tec==5 or com==5: return 5
        result = tec+1 if tec+1 < vp else vp
    elif corso in ('D3','C3','C4','C5'):
        if tec==5 or sen==5 or com==5: return 5
        lim = min(tec,sen)+1
        result = lim if lim < vp else vp
    elif corso == 'D5':
        if tec==5 or sen==5 or com==5: return 5
        result = tec+1 if tec+1 < vp else vp
    elif corso == 'D4':
        if tec==5 or sen==5 or com==5: return 5
        if tec==8 and sen==8: result = max(8, vp)
        else:
            m2 = min(tec,sen)
            result = m2 if min(tec,sen,vp) < m2 else min(tec,sen,vp)
    else:
        return None
    return int(result)

def calcola_finale(pts_list):
    validi = [p for p in pts_list if p is not None]
    if not validi: return None
    import math
    return math.floor(sum(validi)/len(validi))

# ── Auth admin ────────────────────────────────────────────────────────────
def check_admin(f):
    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        token = request.headers.get('X-Admin-Token','')
        with get_db() as db:
            row = db.execute('SELECT token FROM sessions WHERE token=?',(token,)).fetchone()
        if not row:
            return jsonify({'error':'Non autorizzato'}), 401
        return f(*args, **kwargs)
    return wrapper

# ── Routes ────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

@app.route('/api/login', methods=['POST'])
def login():
    data = request.json or {}
    if data.get('password') != ADMIN_PWD:
        return jsonify({'error':'Password errata'}), 401
    token = secrets.token_hex(32)
    with get_db() as db:
        db.execute('INSERT INTO sessions(token) VALUES(?)', (token,))
        db.commit()
    return jsonify({'token': token})

@app.route('/api/valutazioni', methods=['POST'])
def salva_valutazione():
    d = request.json or {}
    corso   = d.get('corso','')
    criteri = ['tec','sen','aff','pro','imp','dis','com']
    giorni  = range(1,8)

    if not d.get('istruttore') or not d.get('allievo') or corso not in PESI:
        return jsonify({'error':'Dati mancanti o non validi'}), 400

    # Raccoglie voti per giorno e calcola punteggi giornalieri
    cols, vals = [], []
    pts = []
    for g in giorni:
        voti_g = [d.get(f'{c}_g{g}') for c in criteri]
        # Converte in int se presenti
        voti_g = [int(v) if v is not None else None for v in voti_g]
        # Valida range 1-10
        for v in voti_g:
            if v is not None and not (1 <= v <= 10):
                return jsonify({'error':f'Voto fuori range (1-10)'}), 400
        pt = calcola_punteggio(corso, voti_g) if all(v is not None for v in voti_g) else None
        pts.append(pt)
        for i,c in enumerate(criteri):
            cols.append(f'{c}_g{g}')
            vals.append(voti_g[i])
        cols.append(f'pts_g{g}')
        vals.append(pt)

    finale = calcola_finale(pts)

    from datetime import date
    fixed_cols = ['data','istruttore','allievo','corso','punteggio_finale']
    fixed_vals = [date.today().isoformat(), d['istruttore'], d['allievo'], corso, finale]

    all_cols = fixed_cols + cols
    all_vals = fixed_vals + vals

    sql = f"INSERT INTO valutazioni ({','.join(all_cols)}) VALUES ({','.join(['?']*len(all_vals))})"
    with get_db() as db:
        db.execute(sql, all_vals)
        db.commit()

    return jsonify({'ok': True, 'punteggio_finale': finale, 'pts_giornalieri': pts})


@app.route("/api/valutazioni/public", methods=["GET"])
def lista_valutazioni_public():
    """Riepilogo visibile a tutti."""
    q     = request.args.get("q","")
    corso = request.args.get("corso","")
    limit = int(request.args.get("limit", 500))
    where, params = [], []
    if q:     where.append("(allievo LIKE ? OR istruttore LIKE ?)"); params += [f"%{q}%", f"%{q}%"]
    if corso: where.append("corso=?"); params.append(corso)
    sql = "SELECT id,data,istruttore,allievo,corso,pts_g1,pts_g2,pts_g3,pts_g4,pts_g5,pts_g6,pts_g7,punteggio_finale FROM valutazioni"
    if where: sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY id DESC LIMIT ?"
    params.append(limit)
    with get_db() as db:
        rows  = db.execute(sql, params).fetchall()
        csql  = "SELECT COUNT(*) FROM valutazioni" + (" WHERE " + " AND ".join(where) if where else "")
        total = db.execute(csql, params[:-1]).fetchone()[0]
    return jsonify({"total": total, "rows": [dict(r) for r in rows]})

@app.route('/api/valutazioni', methods=['GET'])
@check_admin
def lista_valutazioni():
    corso    = request.args.get('corso','')
    allievo  = request.args.get('allievo','')
    istr     = request.args.get('istruttore','')
    limit    = int(request.args.get('limit', 200))
    offset   = int(request.args.get('offset', 0))

    where, params = [], []
    if corso:   where.append('corso=?');       params.append(corso)
    if allievo: where.append('allievo LIKE ?'); params.append(f'%{allievo}%')
    if istr:    where.append('istruttore LIKE ?'); params.append(f'%{istr}%')

    sql = 'SELECT * FROM valutazioni'
    if where: sql += ' WHERE ' + ' AND '.join(where)
    sql += ' ORDER BY id DESC LIMIT ? OFFSET ?'
    params += [limit, offset]

    with get_db() as db:
        rows = db.execute(sql, params).fetchall()
        total = db.execute(
            'SELECT COUNT(*) FROM valutazioni' +
            (' WHERE ' + ' AND '.join(where) if where else ''),
            params[:-2]
        ).fetchone()[0]

    return jsonify({'total': total, 'rows': [dict(r) for r in rows]})

@app.route('/api/valutazioni/<int:vid>', methods=['DELETE'])
@check_admin
def elimina_valutazione(vid):
    with get_db() as db:
        db.execute('DELETE FROM valutazioni WHERE id=?', (vid,))
        db.commit()
    return jsonify({'ok': True})

@app.route('/api/stats', methods=['GET'])
@check_admin
def stats():
    with get_db() as db:
        tot   = db.execute('SELECT COUNT(*) FROM valutazioni').fetchone()[0]
        istr  = db.execute('SELECT COUNT(DISTINCT istruttore) FROM valutazioni').fetchone()[0]
        corsi = db.execute('SELECT COUNT(DISTINCT corso) FROM valutazioni').fetchone()[0]
        media = db.execute('SELECT AVG(punteggio_finale) FROM valutazioni WHERE punteggio_finale IS NOT NULL').fetchone()[0]
        per_corso = db.execute(
            'SELECT corso, COUNT(*) as n, AVG(punteggio_finale) as media FROM valutazioni GROUP BY corso ORDER BY corso'
        ).fetchall()
    return jsonify({
        'totale': tot, 'istruttori': istr, 'corsi_attivi': corsi,
        'media_generale': round(media,1) if media else None,
        'per_corso': [dict(r) for r in per_corso]
    })

@app.route('/api/export/csv', methods=['GET'])
@check_admin
def export_csv():
    import csv, io
    with get_db() as db:
        rows = db.execute('SELECT * FROM valutazioni ORDER BY id').fetchall()
    if not rows:
        return jsonify({'error':'Nessun dato'}), 404

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(rows[0].keys())
    for r in rows:
        writer.writerow(list(r))

    from flask import Response
    return Response(
        '\ufeff' + output.getvalue(),
        mimetype='text/csv',
        headers={'Content-Disposition': 'attachment; filename=CVC_valutazioni.csv'}
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
