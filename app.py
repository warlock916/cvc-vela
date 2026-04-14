from flask import Flask, request, jsonify, send_from_directory, Response
from flask_cors import CORS
import sqlite3, os, secrets, functools, csv, io, math
from datetime import date

app = Flask(__name__, static_folder='static')
CORS(app)

DB        = os.environ.get('DB_PATH', 'cvc.db')
ADMIN_PWD = os.environ.get('ADMIN_PASSWORD', 'admin123')

CRITERI = ['tec','sen','aff','pro','imp','dis','com']
DAY_KEYS = ['g1','g2','g3','g4','g5','g6','fin']

def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    # Colonne voti: tec_g1..tec_fin, sen_g1..sen_fin, ..., com_g1..com_fin
    # Colonne punteggi: pts_g1..pts_fin
    voti_cols  = [f"{c}_{d} INTEGER" for c in CRITERI for d in DAY_KEYS]
    pts_cols   = [f"pts_{d} INTEGER" for d in DAY_KEYS]
    with get_db() as db:
        db.execute(f'''CREATE TABLE IF NOT EXISTS valutazioni (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            data             TEXT NOT NULL,
            istruttore       TEXT NOT NULL,
            corso            TEXT NOT NULL,
            turno            TEXT NOT NULL,
            allievo          TEXT NOT NULL,
            {", ".join(voti_cols)},
            {", ".join(pts_cols)},
            punteggio_finale REAL
        )''')
        db.execute('''CREATE TABLE IF NOT EXISTS sessions (
            token TEXT PRIMARY KEY,
            created_at TEXT DEFAULT (datetime('now'))
        )''')
        db.commit()

init_db()

PESI = {
    'D1':[.30,.10,.10,.15,.12,.13,.10],'D2':[.30,.10,.10,.15,.12,.13,.10],
    'D3':[.30,.15,.15,.10,.08,.12,.10],'D4':[.30,.15,.20,.05,.08,.12,.10],
    'D5':[.30,.15,.24,.05,.08,.08,.10],'C1':[.35,.15,.10,.10,.10,.10,.10],
    'C2':[.35,.15,.10,.10,.10,.10,.10],'C3':[.30,.20,.15,.05,.10,.10,.10],
    'C4':[.25,.25,.15,.05,.10,.10,.10],'C5':[.25,.25,.15,.05,.10,.10,.10],
}

def calcola_punteggio(corso, voti):
    p = PESI.get(corso)
    if not p or None in voti: return None
    tec,sen,aff,pro,imp,dis,com = voti
    vp = sum(v*w for v,w in zip(voti,p))
    if corso in ('D1','D2','C1','C2'):
        if tec==5 or com==5: return 5
        result = tec+1 if tec+1<vp else vp
    elif corso in ('D3','C3','C4','C5'):
        if tec==5 or sen==5 or com==5: return 5
        lim=min(tec,sen)+1; result=lim if lim<vp else vp
    elif corso=='D5':
        if tec==5 or sen==5 or com==5: return 5
        result=tec+1 if tec+1<vp else vp
    elif corso=='D4':
        if tec==5 or sen==5 or com==5: return 5
        if tec==8 and sen==8: result=max(8,vp)
        else:
            m2=min(tec,sen); result=m2 if min(tec,sen,vp)<m2 else min(tec,sen,vp)
    else: return None
    return int(math.floor(result))

def check_admin(f):
    @functools.wraps(f)
    def wrapper(*args,**kwargs):
        token=request.headers.get('X-Admin-Token','')
        with get_db() as db:
            row=db.execute('SELECT token FROM sessions WHERE token=?',(token,)).fetchone()
        if not row: return jsonify({'error':'Non autorizzato'}),401
        return f(*args,**kwargs)
    return wrapper

@app.route('/')
def index():
    return send_from_directory('static','index.html')

@app.route('/api/login', methods=['POST'])
def login():
    d=request.json or {}
    if d.get('password')!=ADMIN_PWD: return jsonify({'error':'Password errata'}),401
    token=secrets.token_hex(32)
    with get_db() as db:
        db.execute('INSERT INTO sessions(token) VALUES(?)',(token,)); db.commit()
    return jsonify({'token':token})

@app.route('/api/scheda', methods=['POST'])
def salva_scheda():
    d       = request.json or {}
    records = d.get('records',[])
    if not records: return jsonify({'error':'Nessun record'}),400

    oggi    = date.today().isoformat()
    salvati = 0

    with get_db() as db:
        for rec in records:
            corso   = rec.get('corso','')
            istr    = rec.get('istruttore','').strip()
            allievo = rec.get('allievo','').strip()
            turno   = rec.get('turno','').strip()
            if not all([corso,istr,allievo,turno]): continue

            cols, vals = [], []

            # Voti per ogni criterio×giorno
            for c in CRITERI:
                for dk in DAY_KEYS:
                    v = rec.get(f'{c}_{dk}')
                    v = int(v) if v is not None and 1<=int(v)<=10 else None
                    cols.append(f'{c}_{dk}'); vals.append(v)

            # Punteggi giornalieri
            pts_list = []
            for dk in DAY_KEYS:
                v = rec.get(f'pts_{dk}')
                pt = int(v) if v is not None else None
                # Ricalcola lato server come verifica
                voti_day = [rec.get(f'{c}_{dk}') for c in CRITERI]
                voti_day = [int(x) if x is not None and 1<=int(x)<=10 else None for x in voti_day]
                pt_srv   = calcola_punteggio(corso, voti_day)
                pt_final = pt_srv if pt_srv is not None else pt
                cols.append(f'pts_{dk}'); vals.append(pt_final)
                pts_list.append(pt_final)

            # Punteggio finale = media punteggi giornalieri
            validi = [p for p in pts_list if p is not None]
            pf     = int(math.floor(sum(validi)/len(validi))) if validi else None

            all_cols = ['data','istruttore','corso','turno','allievo']+cols+['punteggio_finale']
            all_vals = [oggi,istr,corso,turno,allievo]+vals+[pf]
            db.execute(f"INSERT INTO valutazioni ({','.join(all_cols)}) VALUES ({','.join(['?']*len(all_vals))})", all_vals)
            salvati+=1
        db.commit()

    return jsonify({'ok':True,'salvati':salvati})

@app.route('/api/valutazioni/public', methods=['GET'])
def lista_public():
    q     = request.args.get('q','')
    corso = request.args.get('corso','')
    limit = int(request.args.get('limit',500))
    pts_cols = ','.join(f'pts_{dk}' for dk in DAY_KEYS)

    where,params=[],[]
    if q:     where.append('(allievo LIKE ? OR istruttore LIKE ? OR turno LIKE ?)'); params+=[f'%{q}%']*3
    if corso: where.append('corso=?'); params.append(corso)

    sql  = f'SELECT id,data,istruttore,turno,allievo,corso,{pts_cols},punteggio_finale FROM valutazioni'
    csql = 'SELECT COUNT(*) FROM valutazioni'
    if where:
        w=' WHERE '+' AND '.join(where); sql+=w; csql+=w
    sql+=' ORDER BY id DESC LIMIT ?'; params.append(limit)

    with get_db() as db:
        rows  = db.execute(sql,params).fetchall()
        total = db.execute(csql,params[:-1]).fetchone()[0]
    return jsonify({'total':total,'rows':[dict(r) for r in rows]})

@app.route('/api/valutazioni/<int:vid>', methods=['DELETE'])
@check_admin
def elimina(vid):
    with get_db() as db:
        db.execute('DELETE FROM valutazioni WHERE id=?',(vid,)); db.commit()
    return jsonify({'ok':True})

@app.route('/api/stats', methods=['GET'])
@check_admin
def stats():
    with get_db() as db:
        tot  =db.execute('SELECT COUNT(*) FROM valutazioni').fetchone()[0]
        istr =db.execute('SELECT COUNT(DISTINCT istruttore) FROM valutazioni').fetchone()[0]
        corsi=db.execute('SELECT COUNT(DISTINCT corso) FROM valutazioni').fetchone()[0]
        media=db.execute('SELECT AVG(punteggio_finale) FROM valutazioni WHERE punteggio_finale IS NOT NULL').fetchone()[0]
        perc =db.execute('SELECT corso,COUNT(*) n,AVG(punteggio_finale) media FROM valutazioni GROUP BY corso ORDER BY corso').fetchall()
    return jsonify({'totale':tot,'istruttori':istr,'corsi_attivi':corsi,
                    'media_generale':round(media,1) if media else None,
                    'per_corso':[dict(r) for r in perc]})

@app.route('/api/export/csv', methods=['GET'])
@check_admin
def export_csv():
    with get_db() as db:
        rows=db.execute('SELECT * FROM valutazioni ORDER BY id').fetchall()
    if not rows: return jsonify({'error':'Nessun dato'}),404
    out=io.StringIO()
    w=csv.writer(out)
    w.writerow(rows[0].keys())
    for r in rows: w.writerow(list(r))
    return Response('\ufeff'+out.getvalue(),mimetype='text/csv',
                    headers={'Content-Disposition':'attachment; filename=CVC_valutazioni.csv'})

if __name__=='__main__':
    port=int(os.environ.get('PORT',5000))
    app.run(host='0.0.0.0',port=port,debug=False)
