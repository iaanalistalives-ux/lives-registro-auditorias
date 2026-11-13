# -*- coding: utf-8 -*-
"""Carga la base de OPs y la base de partidas desde el Excel del sistema.
Reemplaza a las pantallas ops.html y partidas.html: lee el archivo con las mismas reglas
y llama a las mismas funciones de la base (sync_ops_reemplazar y sync_partidas_reemplazar),
asi que Alertas LIVES, los colores y la fecha de carga siguen funcionando igual.

Uso (lo corre Claude en la PC; el Excel nunca va a GitHub):
  python -I cargar_base.py ops "<ruta del Excel>"                -> solo revisa y muestra el resumen
  python -I cargar_base.py ops "<ruta del Excel>" --comparar     -> ademas dice que cambiaria en la base
  python -I cargar_base.py ops "<ruta del Excel>" --enviar       -> lo deja en espera; Claude lo aplica en la base
  python -I cargar_base.py partidas "<ruta del Excel>" [--comparar] [--enviar]

Reglas (las de las pantallas):
  OPs      : busca las columnas por nombre en TODAS las hojas, de la ultima a la primera. Un archivo sin
             la columna "combo" (el color) se rechaza. Se cargan las filas que tienen OP.
  Partidas : solo la primera pestana, solo filas visibles. Titulos en una de las 10 primeras filas;
             si la fila siguiente tambien es de titulos (la del sistema), se salta.
  En las dos: lo que viene en el archivo se reemplaza y lo que no viene se mantiene.
"""
import sys, json, re, datetime, unicodedata, urllib.request, urllib.error
from collections import Counter
import openpyxl

SB = 'https://xxfufpddgoyhbotjmaiy.supabase.co'
KEY = 'sb_publishable_1g-n6UT70EUlNAo8Lyc04w_jKR4LpXZ'
USUARIO = 'ctalepcio'

COLS_OPS = {
    'op': ['OP', 'NUMOP', 'NROOP', 'CODOP', 'ORDENDEPRODUCCION', 'ORDENPRODUCCION'],
    'grupo': ['GRUPO', 'CODGRUPOPRO', 'CODGRUPOPROD', 'CODGRUPO', 'GRUPOPRO', 'GRUPOPROD'],
    'estilo': ['ESTILOPROPIO', 'ESTPROP', 'ESTILOPROP', 'ESTPROPIO', 'ESTILO'],
    'color': ['COMBO', 'COLORCOMBO', 'COLOR'],
    'po': ['PO', 'NUMPO', 'NROPO', 'CODPO'],
    'codTela': ['CODIGOTELA', 'CODTELA', 'CODIGODETELA', 'CODMATERIAL'],
    'descTela': ['DESCRIPCIONTELA', 'DESCTELA', 'DESCRIPCIONDETELA', 'DESTELA', 'TELA'],
    'destino': ['DESTINO', 'CODDESTINO', 'CODIGODESTINO', 'DESTINOFINAL'],
    'pedido': ['PEDIDO', 'REQUERIDAS', 'REQUERIDA', 'CANTREQUERIDA', 'CANTIDADREQUERIDA', 'CANTIDAD', 'CANT'],
    'fechaCreacion': ['FECHACREACION', 'FECCREACIONOP', 'FECHACREACIONOP', 'FECCREACION', 'FECHADECREACION',
                      'FECHADECREACIONOP', 'FECCREAOP', 'FECHAOP'],
    'cotizacion': ['COTIZACION', 'COTIZ', 'NUMCOTIZACION', 'NROCOTIZACION', 'CODCOTIZACION'],
    'eCliente': ['ESTILOCLIENTE', 'ECLIENTE', 'ESTCLIENTE', 'ESTILODELCLIENTE', 'ESTILOCLI'],
}
COLS_PAR = {
    'partida': ['PARTIDA', 'CODORDPROV', 'ORDPROV', 'NROPARTIDA'],
    'codTela': ['CODIGODETELA', 'CODTELA', 'CODIGOTELA'],
    'nomTela': ['NOMBREDETELA', 'DESTELA', 'DESCTELA', 'NOMBRETELA', 'DESCRIPCIONTELA', 'TELA'],
    'codColor': ['CODIGODECOLOR', 'CODCOLOR', 'CODIGOCOLOR'],
    'nomColor': ['NOMBREDECOLOR', 'DESCOLOR', 'NOMBRECOLOR', 'COLOR'],
    'grupo': ['GRUPO', 'CODGRUPOTEX', 'GRUPOTEX'],
    'peso': ['PESOPROGRAMADO', 'PESOPROGR', 'PESOPROG'],
    'op': ['OP', 'NUMOP', 'NROOP'],
    'oc': ['OC', 'ORDENCORTE', 'ORDENDECORTE', 'NROOC', 'OCORTE'],
}
NOMBRE_PAR = {'partida': 'PARTIDA', 'codTela': 'CODIGO DE TELA', 'nomTela': 'NOMBRE DE TELA',
              'codColor': 'CODIGO DE COLOR', 'nomColor': 'NOMBRE DE COLOR', 'grupo': 'GRUPO',
              'peso': 'PESO PROGRAMADO', 'op': 'OP', 'oc': 'OC'}


# ---------- texto y numeros, igual que en las pantallas ----------
def T(v):
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, float) and v == int(v) and abs(v) < 1e15:
        return str(int(v))                      # 12345.0 -> "12345", como lo escribe el navegador
    if isinstance(v, datetime.datetime):
        return v.strftime('%d/%m/%Y')
    return str(v).strip()


def U(v):
    return T(v).upper()


def norm(s):
    s = unicodedata.normalize('NFD', T(s).upper())
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return re.sub(r'[^A-Z0-9]', '', s)


def num_ops(v):
    """Cantidad de pedido: 0 si no se puede leer (como ops.html)."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return v
    m = re.match(r'-?\d*\.?\d+|-?\d+\.?', re.sub(r'[^0-9.\-]', '', T(v)))
    try:
        n = float(m.group(0)) if m else 0
    except ValueError:
        n = 0
    return int(n) if n == int(n) else n


def num_par(v):
    """Peso programado: None si no se puede leer (como partidas.html)."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return v
    s = re.sub(r'[^0-9.,\-]', '', T(v))
    if not s:
        return None
    s = s.replace(',', '.') if (',' in s and '.' not in s) else s.replace(',', '')
    try:
        return float(s)
    except ValueError:
        return None


def fecha_iso(v):
    def ok(a, m, d):
        try:
            return datetime.date(a, m, d).isoformat()
        except ValueError:
            return ''
    if isinstance(v, datetime.datetime):
        return v.date().isoformat()
    if isinstance(v, datetime.date):
        return v.isoformat()
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 20000 < v < 80000:
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=round(v))).isoformat()
    s = T(v)
    m = re.match(r'^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})', s)
    if m:
        return ok(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        return ok(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return ''


def fecha_txt(iso):
    return iso[8:10] + '/' + iso[5:7] + '/' + iso[0:4] if iso else ''


def abrir(ruta):
    if ruta.lower().endswith('.xls'):
        raise SystemExit('El archivo es .xls (formato antiguo). Guardalo como .xlsx y vuelve a intentar. No se cargo nada.')
    return openpyxl.load_workbook(ruta, data_only=True)


def es_pie(fila):
    """Fila del pie de los reportes con formato LIVES: una sola celda 'Generado el ... fila(s)'."""
    con = [T(v) for v in fila if T(v)]
    return len(con) == 1 and con[0].upper().startswith('GENERADO EL')


def mapear(fila, cols):
    cab = {}
    for i, c in enumerate(fila):
        k = norm(c)
        if k and k not in cab:
            cab[k] = i
    m, faltan = {}, []
    for k, alias in cols.items():
        h = next((cab[a] for a in alias if a in cab), None)
        if h is None:
            faltan.append(k)
        else:
            m[k] = h
    return m, faltan


# ---------- OPs ----------
def leer_ops(ruta):
    wb = abrir(ruta)
    mejor = None
    for ws in reversed(wb.worksheets):
        filas = list(ws.iter_rows(values_only=True))
        if len(filas) < 2:
            continue
        # Titulos: fila 1 en el export del sistema de produccion; fila 3 en el reporte "Consulta de OPs"
        # del sistema de la empresa (formato LIVES, aceptado desde el 13/11/2026: su columna Color es el combo).
        cab, m = -1, None
        for i in range(min(10, len(filas))):
            mm, faltan = mapear(filas[i], COLS_OPS)
            if not faltan:
                cab, m = i, mm
                break
            f = [COLS_OPS[k][0] for k in faltan]
            if mejor is None or len(f) < len(mejor):
                mejor = f
        if cab < 0:
            continue
        ocultas = sum(1 for i in range(cab + 2, len(filas) + 1) if ws.row_dimensions[i].hidden)
        out = []
        for r in filas[cab + 1:]:
            g = lambda k: r[m[k]] if m[k] < len(r) else None
            if not U(g('op')) or es_pie(r):
                continue
            out.append({'op': U(g('op')), 'grupo': U(g('grupo')), 'estilo': U(g('estilo')), 'color': U(g('color')),
                        'po': U(g('po')), 'codTela': U(g('codTela')), 'descTela': T(g('descTela')),
                        'destino': U(g('destino')), 'pedido': num_ops(g('pedido')), 'eCliente': U(g('eCliente')),
                        'fechaCreacion': fecha_iso(g('fechaCreacion')), 'cotizacion': U(g('cotizacion'))})
        if not out:
            raise SystemExit('Ninguna fila tiene OP. No se cargo nada.')
        return out, ws.title, ocultas
    raise SystemExit('Ninguna hoja tiene las columnas necesarias. A la mas parecida le faltan: %s. No se cargo nada.'
                     % ', '.join(mejor or ['todas']))


def filas_ops(F):
    return [[f['op'], f['grupo'], f['estilo'], f['color'], f['po'], f['codTela'], f['descTela'], f['destino'],
             f['pedido'], f['eCliente'], f['fechaCreacion'] or None, f['cotizacion']] for f in F]


def resumen_ops(F, hoja, ocultas):
    fe = [f['fechaCreacion'] for f in F if f['fechaCreacion']]
    print('Hoja usada: %s' % hoja)
    print('Filas: %d | OPs: %d | Grupos: %d' % (len(F), len({f['op'] for f in F}), len({f['grupo'] for f in F if f['grupo']})))
    if fe:
        print('OPs creadas entre el %s y el %s' % (fecha_txt(min(fe)), fecha_txt(max(fe))))
    for cuenta, rotulo in ((sum(1 for f in F if not f['codTela']), 'sin codigo de tela'),
                           (sum(1 for f in F if not f['pedido']), 'sin cantidad de pedido'),
                           (len(F) - len(fe), 'sin fecha de creacion legible'),
                           (sum(1 for f in F if not f['color']), 'sin color (combo)')):
        if cuenta:
            print('AVISO: %d filas %s' % (cuenta, rotulo))
    if ocultas:
        print('AVISO: el archivo tiene %d filas ocultas por filtro. La carga de OPs las incluye (asi lo hacia la pantalla).' % ocultas)


# ---------- Partidas ----------
def leer_partidas(ruta):
    ws = abrir(ruta).worksheets[0]
    ini = ws.min_row
    A = list(ws.iter_rows(min_row=ini, values_only=True))
    cab, m, mejor = -1, None, None
    for r in range(min(10, len(A))):
        mm, faltan = mapear(A[r], COLS_PAR)
        if not faltan:
            cab, m = r, mm
            break
        if mejor is None or len(faltan) < len(mejor):
            mejor = faltan
    if cab < 0:
        raise SystemExit('La primera pestana ("%s") no tiene las columnas necesarias. Le faltan: %s. No se cargo nada.'
                         % (ws.title, ', '.join(NOMBRE_PAR[k] for k in (mejor or NOMBRE_PAR))))
    desde, salto = cab + 1, 0
    if desde < len(A) and not mapear(A[desde], COLS_PAR)[1]:
        desde, salto = desde + 1, 1
    out, filtradas = [], 0
    for i in range(desde, len(A)):
        if ws.row_dimensions[i + ini].hidden:
            filtradas += 1
            continue
        f = A[i]
        g = lambda k: f[m[k]] if m[k] < len(f) else None
        o = {'partida': U(g('partida')), 'codTela': U(g('codTela')), 'nomTela': T(g('nomTela')),
             'codColor': U(g('codColor')), 'nomColor': T(g('nomColor')), 'grupo': U(g('grupo')),
             'peso': num_par(g('peso')), 'op': U(g('op')), 'oc': U(g('oc'))}
        if (not o['partida'] and not o['op'] and not o['oc'] and not o['codTela']) or es_pie(f):
            continue
        out.append(o)
    if not out:
        raise SystemExit('El archivo no trae filas con datos. No se cargo nada.')
    return out, ws.title, salto, filtradas


def filas_partidas(F):
    return [[f['partida'], f['codTela'], f['nomTela'], f['codColor'], f['nomColor'], f['grupo'], f['peso'],
             f['op'], f['oc']] for f in F]


def resumen_partidas(F, hoja, salto, filtradas):
    print('Pestana: %s%s' % (hoja, ' (se salto la segunda fila de titulos, la del sistema)' if salto else ''))
    kg = sum(f['peso'] for f in F if f['peso'] is not None)
    print('Filas: %d | Partidas: %d | OP: %d | Ordenes de corte: %d | Grupos: %d | Kg programados: %.3f' % (
        len(F), len({f['partida'] for f in F if f['partida']}), len({f['op'] for f in F if f['op']}),
        len({f['oc'] for f in F if f['oc']}), len({f['grupo'] for f in F if f['grupo']}), kg))
    rep = Counter(tuple(f[k] for k in ('partida', 'codTela', 'codColor', 'grupo', 'peso', 'op', 'oc')) for f in F)
    for cuenta, rotulo in ((sum(1 for f in F if not f['partida']), 'sin partida (la base no las guarda)'),
                           (sum(1 for f in F if not f['op']), 'sin OP'), (sum(1 for f in F if not f['oc']), 'sin OC'),
                           (sum(1 for f in F if f['peso'] is None), 'sin peso legible'),
                           (sum(c - 1 for c in rep.values() if c > 1), 'repetidas (iguales a otra fila del archivo)')):
        if cuenta:
            print('AVISO: %d filas %s' % (cuenta, rotulo))
    if filtradas:
        print('AVISO: %d filas ocultas por filtro no se usan (regla: solo filas visibles).' % filtradas)


# ---------- la base ----------
def pedir(ruta, cuerpo=None, metodo='POST', espera=180, extra=None, quiero=None):
    cab = {'Content-Type': 'application/json', 'apikey': KEY}
    cab.update(extra or {})
    if quiero:
        ruta += '?select=' + quiero
    datos = None if cuerpo is None else json.dumps(cuerpo).encode('utf-8')
    req = urllib.request.Request(SB + ruta, data=datos, method=metodo, headers=cab)
    try:
        with urllib.request.urlopen(req, timeout=espera) as r:
            return json.loads(r.read().decode('utf-8') or 'null')
    except urllib.error.HTTPError as e:
        raise SystemExit('La base respondio %s: %s' % (e.code, e.read().decode('utf-8', 'replace')[:300]))
    except urllib.error.URLError as e:
        raise SystemExit('Sin conexion con la base de datos (%s). La tabla quedo como estaba.' % e.reason)


def dejar_en_espera(que, archivo, filas):
    """OPs: la base solo acepta la carga con la sesion de Christian, y el puente de sesion no se la da a un
    programa (a proposito). Por eso el archivo ya leido se deja en la tabla de espera `sync_carga_archivo`
    y Claude lo aplica desde su conexion a la base con la misma funcion sync_ops_reemplazar.
    PENDIENTE: la tabla de espera se crea solo con el "procede" de Christian."""
    r = pedir('/rest/v1/sync_carga_archivo', {'tabla': que, 'archivo': archivo, 'filas': filas, 'usuario': USUARIO},
              extra={'Prefer': 'return=representation'}, quiero='id,tabla,archivo,creado')
    return r[0] if isinstance(r, list) and r else r


def tabla(nombre, columnas):
    out, paso = [], 1000
    while True:
        lote = pedir('/rest/v1/%s?select=%s&order=id&limit=%d&offset=%d' % (nombre, ','.join(columnas), paso, len(out)),
                     metodo='GET', espera=60)
        out += lote
        if len(lote) < paso:
            return out


def _n(v):
    return None if v is None else round(float(v), 6)


def comparar(que, filas):
    """Compara lo que trae el archivo con lo que hay hoy en la base, sin cambiar nada."""
    if que == 'ops':
        cols = ['op', 'grupo', 'estilo', 'color', 'po', 'cod_tela', 'desc_tela', 'destino', 'pedido', 'e_cliente',
                'fecha_creacion', 'cotizacion']
        nuevo = [tuple([f[0], f[1], f[2], f[3], f[4], f[5], f[6], f[7], _n(f[8]), f[9], f[10], f[11]]) for f in filas]
        viejo = [tuple([r['op'] or '', r['grupo'] or '', r['estilo'] or '', r['color'] or '', r['po'] or '',
                        r['cod_tela'] or '', r['desc_tela'] or '', r['destino'] or '', _n(r['pedido'] or 0),
                        r['e_cliente'] or '', r['fecha_creacion'], r['cotizacion'] or '']) for r in tabla('ops', cols)]
        cosa = 'OPs'
    else:
        cols = ['partida', 'codigo_tela', 'nombre_tela', 'codigo_color', 'nombre_color', 'grupo', 'peso_programado',
                'op', 'oc']
        nuevo = [tuple([f[0], f[1], f[2] or None, f[3], f[4] or None, f[5], _n(f[6]), f[7], f[8]]) for f in filas if f[0]]
        viejo = [tuple([r['partida'], r['codigo_tela'] or '', r['nombre_tela'], r['codigo_color'] or '',
                        r['nombre_color'], r['grupo'] or '', _n(r['peso_programado']), r['op'] or '', r['oc'] or ''])
                 for r in tabla('partidas', cols)]
        cosa = 'partidas'
    en_archivo = {f[0] for f in nuevo}
    en_base = {r[0] for r in viejo}
    cn, cv = Counter(nuevo), Counter(r for r in viejo if r[0] in en_archivo)
    cambian = {t[0] for t in (cn - cv)} | {t[0] for t in (cv - cn)}
    nuevas = en_archivo - en_base
    print('--- Comparacion con la base (no se cambio nada) ---')
    print('Hoy en la base: %d filas, %d %s' % (len(viejo), len(en_base), cosa))
    print('%s del archivo: %d | nuevas: %d | con algun cambio: %d | iguales: %d | se mantienen sin tocar: %d' % (
        cosa, len(en_archivo), len(nuevas), len(cambian - nuevas), len(en_archivo) - len(cambian | nuevas),
        len(en_base - en_archivo)))
    for rotulo, lista in (('Nuevas', nuevas), ('Con cambio', cambian - nuevas)):
        if lista:
            print('%s (hasta 15): %s' % (rotulo, ', '.join(sorted(lista)[:15])))


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    if len(sys.argv) < 3 or sys.argv[1] not in ('ops', 'partidas'):
        raise SystemExit(__doc__)
    que, ruta = sys.argv[1], sys.argv[2]
    print('Archivo: %s' % ruta.replace('\\', '/').split('/')[-1])
    if que == 'ops':
        F, hoja, ocultas = leer_ops(ruta)
        resumen_ops(F, hoja, ocultas)
        filas, funcion = filas_ops(F), 'sync_ops_reemplazar'
    else:
        F, hoja, salto, filtradas = leer_partidas(ruta)
        resumen_partidas(F, hoja, salto, filtradas)
        filas, funcion = filas_partidas(F), 'sync_partidas_reemplazar'
    if '--json' in sys.argv:
        sal = sys.argv[sys.argv.index('--json') + 1]
        with open(sal, 'w', encoding='utf-8') as f:
            json.dump(filas, f, ensure_ascii=False)
        print('JSON guardado en', sal)
    if '--comparar' in sys.argv:
        comparar(que, filas)
    if '--enviar' in sys.argv and que == 'ops':
        r = dejar_en_espera(que, ruta.replace('\\', '/').split('/')[-1], filas)
        print('En espera:', json.dumps(r, ensure_ascii=False))
        print('FALTA APLICAR: las %d filas quedaron en espera. Claude las aplica con sync_ops_reemplazar y revisa la respuesta.' % len(filas))
        return
    if '--enviar' in sys.argv:
        r = pedir('/rest/v1/rpc/' + funcion, {'p_filas': filas, 'p_usuario': USUARIO.upper(), 'p_forzar': False})
        print('Respuesta de la base:', json.dumps(r, ensure_ascii=False))
        if not (r and r.get('ok')):
            raise SystemExit('NO SE CARGO: %s' % ((r or {}).get('error') or 'la base no confirmo la carga'))
        esperado = len(filas) if que == 'ops' else sum(1 for f in filas if f[0])
        if r.get('cargadas') != esperado:
            raise SystemExit('ATENCION: se leyeron %d filas pero se cargaron %s. Revisar antes de seguir.' % (esperado, r.get('cargadas')))
        print('Listo: %d filas cargadas. La tabla tiene %s filas en total (%s que el archivo no traia se mantuvieron).'
              % (r['cargadas'], r.get('total'), r.get('mantenidas')))
        if que == 'ops':
            if r.get('cambios_error') or not (r.get('cambios') is not None and r['cambios'] >= 0):
                print('ATENCION: las OPs si se cargaron, pero la comparacion para Alertas fallo (%s). Esta carga no genero alertas.' % r.get('cambios_error'))
            else:
                print('Alertas: %s cambio(s) detectados.' % r['cambios'])
            if r.get('colores_propagados_error'):
                print('ATENCION: fallo la propagacion de colores (%s).' % r['colores_propagados_error'])


if __name__ == '__main__':
    main()
