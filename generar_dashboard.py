# -*- coding: utf-8 -*-
"""
=============================================================================
 GENERADOR AUTOMÁTICO · Dashboard de Reabastecimiento ABC · Cliente GOO
=============================================================================
 Qué hace:
   1. Lee los dos archivos Excel (inventario completo + pendiente por pick).
   2. Recalcula TODAS las métricas con la lógica del proyecto.
   3. Genera dos archivos HTML:
        - Dashboard_Reabastecimiento_TipoA_GOO.html  (interactivo)
        - Informe_Reabastecimiento_GOO_correo.html   (estático, para correo)
   4. Guarda el corte actual para comparar el avance la próxima vez.

 Cómo se usa:
   - Coloca este script junto a la plantilla y los Excel (ver nombres abajo).
   - Doble clic en "Generar_Dashboard.bat" (Windows).

 Archivos que espera encontrar en la MISMA carpeta:
   - plantilla_dashboard.html               (la plantilla, no borrar)
   - Para_analisis_de_GOO_proyecto.xlsx     (inventario completo)
   - Analisis_de_progreso_ABC_GOO.xlsx      (pendiente por pick)
=============================================================================
"""
import sys, os, json, re, math, html as htmlmod, datetime

# ---- Localizar carpeta del script (funciona con doble clic) ----
CARPETA = os.path.dirname(os.path.abspath(__file__))
def ruta(nombre): return os.path.join(CARPETA, nombre)

# Nombres de archivo configurables
ARCH_INVENTARIO = "Para_analisis_de_GOO_proyecto.xlsx"
ARCH_PENDIENTE  = "Analisis_de_progreso_ABC_GOO.xlsx"
HOJA_PENDIENTE_OPCIONES = ["Pendiente por pick", "Pendiente por Pick", "Pendiente por PICK"]
PLANTILLA       = "plantilla_dashboard.html"
PLANTILLA_HERRAMIENTA = "plantilla_herramienta.html"
SALIDA_INTERACTIVO = "Dashboard_Reabastecimiento_TipoA_GOO.html"
SALIDA_CORREO      = "Informe_Reabastecimiento_GOO_correo.html"
SALIDA_HERRAMIENTA = "Herramienta_Reabastecimiento_GOO.html"
ARCH_CORTE_ANTERIOR = "corte_anterior.json"  # memoria del avance

def aviso(msg): print("  " + msg)

# ---- Verificar dependencias ----
try:
    import pandas as pd
except ImportError:
    print("\n[ERROR] Falta la librería 'pandas'. Instálala abriendo una terminal y escribiendo:")
    print("        pip install pandas openpyxl\n")
    input("Presiona ENTER para cerrar...")
    sys.exit(1)

print("="*70)
print("  GENERANDO DASHBOARD DE REABASTECIMIENTO · CLIENTE GOO")
print("="*70)

# ---- Verificar que existan los archivos ----
faltan = []
for f in [ARCH_INVENTARIO, ARCH_PENDIENTE, PLANTILLA]:
    if not os.path.exists(ruta(f)):
        faltan.append(f)
if faltan:
    print("\n[ERROR] No encontré estos archivos en la carpeta:")
    for f in faltan: print("        - " + f)
    print("\n  Asegúrate de que estén junto a este programa.\n")
    input("Presiona ENTER para cerrar...")
    sys.exit(1)

# =====================================================================
#  LECTURA DE DATOS
# =====================================================================
aviso("Leyendo inventario...")
inv = pd.read_excel(ruta(ARCH_INVENTARIO))
inv['Area'] = inv['Area'].astype(str).str.strip()
inv['Tipo'] = inv['Tipo (ABC)'].astype(str).str.strip()
inv['CID']  = inv['CID'].astype(str)

# -------------------------------------------------------------------
# DEFINICIÓN DE "PISO":
#   - Área MEZZANINE BINS, O BIEN
#   - Ubicación que empieza con MZ06 y termina en -01
#     (físicamente ya están a nivel de piso aunque el sistema diga RACK).
# -------------------------------------------------------------------
def es_piso(area, ubicacion):
    if str(area).strip() == 'MEZZANINE BINS':
        return True
    u = str(ubicacion).strip().upper()
    if u.startswith('MZ06') and u.endswith('-01'):
        return True
    return False

inv['Ubicación'] = inv['Ubicación'].astype(str).str.strip()
inv['en_piso'] = inv.apply(lambda r: es_piso(r['Area'], r['Ubicación']), axis=1)

aviso("Leyendo pendiente por pick...")
xl = pd.ExcelFile(ruta(ARCH_PENDIENTE))
hoja = next((h for h in HOJA_PENDIENTE_OPCIONES if h in xl.sheet_names), xl.sheet_names[0])
pend = pd.read_excel(xl, sheet_name=hoja)
pend['ABC']  = pend['Tipo (ABC)'].astype(str).str.strip()
pend['Area'] = pend['Area'].astype(str).str.strip()
pend['Ubic'] = pend['Ubic'].astype(str).str.strip()
# PISO = MEZZANINE BINS o MZ06...-01 ; ALTURA = todo lo demás
pend['loc']  = pend.apply(lambda r: 'piso' if es_piso(r['Area'], r['Ubic']) else 'altura', axis=1)

# =====================================================================
#  CÁLCULOS · PENDIENTE
# =====================================================================
aviso("Calculando métricas del pendiente...")
total_und = int(pend['Pendiente'].sum())
matriz = {t: {'piso':  int(pend[(pend['ABC']==t)&(pend['loc']=='piso')]['Pendiente'].sum()),
              'altura':int(pend[(pend['ABC']==t)&(pend['loc']=='altura')]['Pendiente'].sum())}
          for t in ['A','B','C']}
piso_und   = int(pend[pend['loc']=='piso']['Pendiente'].sum())
altura_und = int(pend[pend['loc']=='altura']['Pendiente'].sum())
a_total  = int(pend[pend['ABC']=='A']['Pendiente'].sum())
a_altura = matriz['A']['altura']
a_piso   = matriz['A']['piso']
a_alt_df = pend[(pend['ABC']=='A')&(pend['loc']=='altura')]

por_pedido = []
for pedido, g in pend.groupby('Pedido'):
    por_pedido.append({
        'pedido': str(pedido).split(' - ')[0].split(' -')[0].strip(),
        'nombre': str(pedido),
        'total':  int(g['Pendiente'].sum()),
        'a_altura': int(g[(g['ABC']=='A')&(g['loc']=='altura')]['Pendiente'].sum())})
por_pedido = sorted(por_pedido, key=lambda x: -x['total'])

descol = 'Descripción' if 'Descripción' in pend.columns else None
top = a_alt_df.groupby(['Producto'] + ([descol] if descol else []) + ['Ubic'])['Pendiente'].sum() \
              .reset_index().sort_values('Pendiente', ascending=False).head(12)
top_a = [{'sku':r['Producto'],'desc':(r[descol] if descol else ''),'ubic':r['Ubic'],'und':int(r['Pendiente'])}
         for _, r in top.iterrows()]
por_pais = {str(k):int(v) for k,v in pend.groupby('Pais')['Pendiente'].sum().items()} if 'Pais' in pend.columns else {}
catcol = 'Categoría' if 'Categoría' in pend.columns else ('Categoria' if 'Categoria' in pend.columns else None)
por_cat = {str(k):int(v) for k,v in pend.groupby(catcol)['Pendiente'].sum().items()} if catcol else {}

# Cobertura de piso por OID
por_oid = []
oidcol = 'OID' if 'OID' in pend.columns else None
if oidcol:
    for oid, g in pend.groupby(oidcol):
        tt = int(g['Pendiente'].sum()); pp = int(g[g['loc']=='piso']['Pendiente'].sum())
        por_oid.append({'oid':int(oid),'pedido':str(g['Pedido'].iloc[0]),
                        'total':tt,'piso':pp,'altura':tt-pp,
                        'pct': round(100*pp/tt,1) if tt else 0})
    por_oid = sorted(por_oid, key=lambda x: -x['pct'])

# =====================================================================
#  CÁLCULOS · REABASTECIMIENTO (cajas a BINS)
# =====================================================================
aviso("Calculando reabastecimiento a piso...")
req = {'A':2, 'B':1}
RB = {}
for tipo in ['A','B']:
    sub = inv[inv['Tipo']==tipo]; needed = req[tipo]
    faltan_c = afect = cero = ok = 0
    for prod, g in sub.groupby('Producto'):
        have  = g[g['en_piso']]['CID'].nunique()
        falta = max(0, needed - have)
        if falta > 0:
            avail = g[(~g['en_piso']) & (g['Area']=='MEZZANINE RACK')]['CID'].nunique()
            faltan_c += min(falta, avail); afect += 1
            if have == 0: cero += 1
        else:
            ok += 1
    RB[tipo] = dict(faltan=int(faltan_c), incompletos=int(afect), sin_caja=int(cero),
                    ok=int(ok), productos=int(sub['Producto'].nunique()), req=needed)
subC = inv[inv['Tipo']=='C']
cout = subC[subC['Area']!='MEZZANINE RACK']
RB['C'] = dict(fuera_rack=int(cout['CID'].nunique()),
               ok=int(subC['Producto'].nunique() - cout['Producto'].nunique()),
               productos=int(subC['Producto'].nunique()),
               bins=int(cout[cout['Area']=='MEZZANINE BINS']['CID'].nunique()),
               paleta=int(cout[cout['Area']=='Paleta']['CID'].nunique()),
               otro=int(cout['CID'].nunique() - cout[cout['Area']=='MEZZANINE BINS']['CID'].nunique()
                        - cout[cout['Area']=='Paleta']['CID'].nunique()))
RB['total'] = RB['A']['faltan'] + RB['B']['faltan']

def cajas_faltan(tipo, needed):
    sub = inv[inv['Tipo']==tipo]; tot = 0
    for prod, g in sub.groupby('Producto'):
        have  = g[g['en_piso']]['CID'].nunique()
        avail = g[(~g['en_piso']) & (g['Area']=='MEZZANINE RACK')]['CID'].nunique()
        tot += min(max(0, needed-have), avail)
    return tot
cov = dict(A=dict(ok=RB['A']['ok'], comp=RB['A']['incompletos'], cajas=cajas_faltan('A',2)),
           B=dict(ok=RB['B']['ok'], comp=RB['B']['incompletos'], cajas=cajas_faltan('B',1)),
           C=dict(ok=RB['C']['ok'], comp=int(cout['Producto'].nunique()), cajas=cout['CID'].nunique()))

# B que baja a piso vs. lo que queda
inv_b = inv[(inv['Tipo']=='B')&(inv['Area']=='MEZZANINE RACK')]
maxbox = inv_b.groupby('Producto')['Saldo'].max()
b_alt = pend[(pend['ABC']=='B')&(pend['loc']=='altura')].groupby('Producto')['Pendiente'].sum()
b_baja = b_queda = 0
for prod, u in b_alt.items():
    cap = maxbox[prod] if prod in maxbox.index else 0
    bb = min(u, cap); b_baja += bb; b_queda += max(0, u-bb)
b_baja = int(b_baja); b_queda = int(b_queda)

# Recorridos en altura tras la mejora (C-altura + B remanente)
c_alt_df = pend[(pend['ABC']=='C')&(pend['loc']=='altura')]
b_alt_df = pend[(pend['ABC']=='B')&(pend['loc']=='altura')]
brem_ubic = set()
for prod, g in b_alt_df.groupby('Producto'):
    u = int(g['Pendiente'].sum()); cap = maxbox[prod] if prod in maxbox.index else 0
    if max(0, u-cap) > 0: brem_ubic.update(g['Ubic'].unique())
ubic_altura_after = set(c_alt_df['Ubic'].unique()) | brem_ubic
recorridos_after = len(ubic_altura_after)
altura_after = b_queda + matriz['C']['altura']
rec_prom = round(altura_after/recorridos_after, 1) if recorridos_after else 0

# Corridas (estilo+color = base del SKU sin la última talla)
def base_sku(s):
    s = str(s); p = s.rsplit('-', 1); return p[0] if len(p)==2 else s
pend['base'] = pend['Producto'].apply(base_sku)
g_corr = pend.groupby('base')['Ubic'].nunique()
multi = g_corr[g_corr > 1]
corr = dict(corridas=int(g_corr.shape[0]), multi=int(multi.shape[0]),
            prom=round(float(multi.mean()),1) if len(multi) else 0,
            maxv=int(g_corr.max()), visits=int(g_corr.sum()), cons=int(g_corr.shape[0]),
            pct=round(100*(int(g_corr.sum())-g_corr.shape[0])/int(g_corr.sum()),1) if g_corr.sum() else 0)
masd = g_corr.sort_values(ascending=False).index[0]
corr['desc'] = str(pend[pend['base']==masd]['Descripción'].iloc[0]) if 'Descripción' in pend.columns else ''

# % cumplimiento del estándar
def cumplimiento(tipo, needed):
    sub = inv[inv['Tipo']==tipo]; rt = eb = 0
    for prod, g in sub.groupby('Producto'):
        have  = g[g['en_piso']]['CID'].nunique()
        avail = g[(~g['en_piso']) & (g['Area']=='MEZZANINE RACK')]['CID'].nunique()
        alc = min(needed, have+avail); rt += alc; eb += min(have, alc)
    return rt, eb
rtA, ebA = cumplimiento('A',2); rtB, ebB = cumplimiento('B',1)
cumpl = dict(A=round(100*ebA/rtA,1) if rtA else 0, B=round(100*ebB/rtB,1) if rtB else 0,
             total=round(100*(ebA+ebB)/(rtA+rtB),1) if (rtA+rtB) else 0,
             req=rtA+rtB, en_bins=ebA+ebB, faltan=(rtA+rtB)-(ebA+ebB))

# Concentración 80% del A en altura
pp_alt = sorted([p for p in por_pedido if p['a_altura']>0], key=lambda x:-x['a_altura'])
tot_aalt = sum(p['a_altura'] for p in pp_alt); acum = conc_n = 0
for p in pp_alt:
    acum += p['a_altura']; conc_n += 1
    if acum >= 0.8*tot_aalt: break
top3 = pp_alt[:3]
conc_top3_pct = round(100*sum(x['a_altura'] for x in top3)/tot_aalt) if tot_aalt else 0
def nombre_corto(nm): return re.sub(r'^GOO\d+\s*-?\s*','',nm).strip() or nm
conc_top3 = ', '.join(nombre_corto(p['nombre']) for p in top3)

# Objeto S para el dashboard interactivo
S = dict(total_und=total_und, total_lineas=len(pend), n_pedidos=int(pend['Pedido'].nunique()),
         n_paises=int(pend['Pais'].nunique()) if 'Pais' in pend.columns else 0,
         abc_und={t:int(pend[pend['ABC']==t]['Pendiente'].sum()) for t in ['A','B','C']},
         abc_lineas={t:int((pend['ABC']==t).sum()) for t in ['A','B','C']},
         piso_und=piso_und, altura_und=altura_und, a_total=a_total, a_altura=a_altura, a_piso=a_piso,
         a_altura_skus=int(a_alt_df['Producto'].nunique()), a_altura_pedidos=int(a_alt_df['Pedido'].nunique()),
         matriz=matriz, por_pedido=por_pedido, por_pais=por_pais, por_categoria=por_cat,
         top_a_altura=top_a,
         reabasto={'total':RB['total'],
                   'A':{k:RB['A'][k] for k in ['faltan','incompletos','sin_caja','ok','productos','req']},
                   'B':{k:RB['B'][k] for k in ['faltan','incompletos','ok','productos','req']},
                   'C':{'fuera_rack':RB['C']['fuera_rack'],'ok':RB['C']['ok'],'productos':RB['C']['productos']}},
         b_baja_piso=b_baja, b_queda_altura=b_queda, por_oid=por_oid)

# Guardar resultados intermedios para la parte 2 (generación de HTML)
contexto = dict(S=S, RB=RB, cov=cov, cumpl=cumpl, corr=corr,
                inv_lineas=len(inv), altura_after=altura_after,
                recorridos_after=recorridos_after, rec_prom=rec_prom,
                c_det=dict(bins=RB['C']['bins'], paleta=RB['C']['paleta'], otro=RB['C']['otro']),
                conc_n=conc_n, conc_top3=conc_top3, conc_top3_pct=conc_top3_pct)
with open(ruta('_contexto_tmp.json'), 'w', encoding='utf-8') as f:
    json.dump(contexto, f, ensure_ascii=False)

aviso("Métricas calculadas correctamente.")

# =====================================================================
#  AVANCE · comparar con el corte anterior
# =====================================================================
aviso("Comparando con el corte anterior...")
act = dict(A=RB['A']['faltan'], B=RB['B']['faltan'], C=RB['C']['fuera_rack'], total=RB['total'])
if os.path.exists(ruta(ARCH_CORTE_ANTERIOR)):
    with open(ruta(ARCH_CORTE_ANTERIOR), encoding='utf-8') as f:
        ant = json.load(f)
    primera_vez = False
else:
    # Primera vez: no hay con qué comparar, usamos el mismo corte (deltas en 0)
    ant = dict(act)
    primera_vez = True

def fmt(x): return f"{int(round(x)):,}".replace(",", ".")
def num_es(x):  # número con coma decimal estilo español
    s = f"{x:.1f}".rstrip('0').rstrip('.')
    return s.replace('.', ',')
def signo(x): return ('+' if x > 0 else ('−' if x < 0 else '')) + str(abs(int(x)))
def signo_pct(a, b):
    p = round(100*(b-a)/a, 1) if a else 0
    s = num_es(abs(p))
    return ('+' if p > 0 else ('−' if p < 0 else '')) + s + '%'
def clase(dif, bueno_baja=True):
    # bueno_baja: bajar es bueno (verde 'pos'); subir es malo (rojo 'neg')
    if dif == 0: return ''
    if bueno_baja: return 'pos' if dif < 0 else 'neg'
    return 'neg' if dif < 0 else 'pos'

def avance_de(tipo, bueno_baja=True):
    a = ant[tipo]; b = act[tipo]; dif = b - a
    return dict(ant=fmt(a), act=fmt(b), dif=signo(dif),
                pct=signo_pct(a, b), cls=clase(dif, bueno_baja),
                delta=f"{signo(dif)} cajas · {signo_pct(a,b)}")

avA = avance_de('A'); avB = avance_de('B'); avC = avance_de('C', bueno_baja=True); avT = avance_de('total')

# ---- PROYECCIÓN DE TIEMPO AL 100% ----
# Ritmo real = cajas A+B bajadas entre cortes ÷ días transcurridos.
hoy = datetime.date.today()
proy = {'tiene': False}
faltan_actual = act['total']  # cajas A+B que todavía faltan por bajar
ant_fecha_str = ant.get('fecha') if isinstance(ant, dict) else None
if (not primera_vez) and ant_fecha_str:
    try:
        ant_fecha = datetime.date.fromisoformat(ant_fecha_str)
        dias = (hoy - ant_fecha).days
        bajadas = ant['total'] - act['total']   # cuántas cajas bajaron desde el corte anterior
        if dias > 0 and bajadas > 0:
            ritmo = bajadas / dias              # cajas por día
            dias_faltan = math.ceil(faltan_actual / ritmo) if ritmo > 0 else 0
            fecha_fin = hoy + datetime.timedelta(days=dias_faltan)
            meses = dias_faltan / 30.0
            semanas = dias_faltan / 7.0
            proy = {'tiene': True,
                    'ritmo': round(ritmo, 1),
                    'dias_faltan': int(dias_faltan),
                    'semanas': round(semanas, 1),
                    'meses': round(meses, 1),
                    'fecha_fin': fecha_fin.strftime('%d/%m/%Y'),
                    'bajadas': int(bajadas),
                    'dias_periodo': int(dias),
                    'faltan': int(faltan_actual)}
    except (ValueError, TypeError):
        proy = {'tiene': False}

# Guardar el corte ACTUAL (con fecha) como "anterior" para la próxima corrida
with open(ruta(ARCH_CORTE_ANTERIOR), 'w', encoding='utf-8') as f:
    json.dump(dict(act, fecha=hoy.isoformat()), f, ensure_ascii=False)

# =====================================================================
#  RELLENAR LA PLANTILLA (dashboard interactivo)
# =====================================================================
aviso("Generando dashboard interactivo...")
with open(ruta(PLANTILLA), encoding='utf-8') as f:
    plantilla = f.read()

pa = piso_und + a_altura + b_baja      # piso después de la mejora
pct_after = round(100*pa/total_und) if total_und else 0
hoy_pct_dec = round(100*piso_und/total_und, 1) if total_und else 0

# ---- HTML del bloque de proyección ----
if proy['tiene']:
    if proy['meses'] >= 1.5:
        horizonte = f"{num_es(proy['meses'])} meses"
    elif proy['semanas'] >= 1.5:
        horizonte = f"{num_es(proy['semanas'])} semanas"
    else:
        horizonte = f"{proy['dias_faltan']} días"
    proy_bloque = (
        '<div class="proy">'
        f'<div class="p-num">{horizonte}<small>al ritmo actual</small></div>'
        '<div class="p-body">'
        f'<div class="p-lead">Al paso que vamos, el reabastecimiento llegaría al <b>100%</b> alrededor del <b>{proy["fecha_fin"]}</b>.</div>'
        '<div class="p-stats">'
        f'<span class="ps"><b>{fmt(proy["ritmo"])}</b> cajas por día (ritmo medido)</span>'
        f'<span class="ps"><b>{fmt(proy["faltan"])}</b> cajas A+B por bajar</span>'
        f'<span class="ps"><b>{fmt(proy["bajadas"])}</b> bajadas en los últimos {proy["dias_periodo"]} días</span>'
        '</div></div></div>')
else:
    proy_bloque = (
        '<div class="proy-wait"><span class="pw-ico">⧗</span>'
        '<div class="pw-txt"><b>Proyección de tiempo:</b> se calculará automáticamente cuando exista '
        'un segundo corte con avance. El generador guarda la fecha de cada corrida; en la próxima '
        'que cargues datos nuevos, medirá el ritmo real (cajas bajadas ÷ días transcurridos) y '
        'estimará la fecha de llegada al 100%.</div></div>')

valores = {
  '__DATOS_JSON__': json.dumps(S, ensure_ascii=False),
  '__PROY_BLOQUE__': proy_bloque,
  '__HERO_PCT_A__': str(round(100*S['abc_und']['A']/total_und) if total_und else 0),
  '__INV_LINEAS__': fmt(len(inv)),
  '__RA_INC__': fmt(RB['A']['incompletos']), '__RA_SIN__': fmt(RB['A']['sin_caja']), '__RA_OK__': fmt(RB['A']['ok']),
  '__RB_INC__': fmt(RB['B']['incompletos']), '__RB_OK__': fmt(RB['B']['ok']),
  '__RC_BINS__': fmt(RB['C']['bins']), '__RC_PAL__': fmt(RB['C']['paleta']),
  '__RC_OTRO__': fmt(RB['C']['otro']), '__RC_OK__': fmt(RB['C']['ok']),
  '__COVA_OK__': fmt(cov['A']['ok']), '__COVA_COMP__': fmt(cov['A']['comp']), '__COVA_CAJ__': fmt(cov['A']['cajas']),
  '__COVB_OK__': fmt(cov['B']['ok']), '__COVB_COMP__': fmt(cov['B']['comp']), '__COVB_CAJ__': fmt(cov['B']['cajas']),
  '__COVC_OK__': fmt(cov['C']['ok']), '__COVC_COMP__': fmt(cov['C']['comp']), '__COVC_CAJ__': fmt(cov['C']['cajas']),
  '__COVT_OK__': fmt(cov['A']['ok']+cov['B']['ok']+cov['C']['ok']),
  '__COVT_COMP__': fmt(cov['A']['comp']+cov['B']['comp']+cov['C']['comp']),
  '__COVT_CAJ__': fmt(cov['A']['cajas']+cov['B']['cajas']+cov['C']['cajas']),
  '__HOY_PCT__': str(round(hoy_pct_dec)), '__HOY_PCT_DEC__': str(hoy_pct_dec).replace('.', '.'),
  '__AFTER_PCT__': str(pct_after),
  '__ALT_HOY__': fmt(altura_und), '__ALT_AFTER__': fmt(altura_after),
  '__B_QUEDA__': fmt(b_queda), '__C_ALT__': fmt(matriz['C']['altura']),
  '__REC_N__': str(recorridos_after), '__REC_PROM__': num_es(rec_prom),
  '__N_PEDIDOS__': str(S['n_pedidos']),
  '__A_ALT__': fmt(a_altura), '__A_TOTAL__': fmt(a_total),
  '__A_DESTRABA__': num_es(round(100*a_altura/total_und,1)) if total_und else '0',
  '__PISO_HOY__': fmt(piso_und), '__PISO_AFTER__': fmt(pa), '__B_BAJA__': fmt(b_baja),
  '__TOTAL_LINEAS__': fmt(len(pend)), '__TOTAL_UND__': fmt(total_und),
  # Avance
  '__AV_A_ANT__': avA['ant'], '__AV_A_ACT__': avA['act'], '__AV_A_CLS__': avA['cls'],
  '__AV_A_DELTA__': avA['delta'], '__AV_A_DIF__': avA['dif'], '__AV_A_PCT__': avA['pct'],
  '__AV_B_ANT__': avB['ant'], '__AV_B_ACT__': avB['act'], '__AV_B_CLS__': avB['cls'],
  '__AV_B_DELTA__': avB['delta'], '__AV_B_DIF__': avB['dif'], '__AV_B_PCT__': avB['pct'],
  '__AV_C_ANT__': avC['ant'], '__AV_C_ACT__': avC['act'], '__AV_C_CLS__': avC['cls'],
  '__AV_C_DIF__': avC['dif'], '__AV_C_PCT__': avC['pct'],
  '__AV_T_ANT__': avT['ant'], '__AV_T_ACT__': avT['act'], '__AV_T_CLS__': avT['cls'],
  '__AV_T_DELTA__': avT['delta'], '__AV_T_DIF__': avT['dif'], '__AV_T_PCT__': avT['pct'],
  '__AV_T_DIFABS__': str(abs(act['total']-ant['total'])),
  # Cumplimiento
  '__CUMPL_TOTAL__': num_es(cumpl['total']), '__CUMPL_TOTAL_DEC__': str(cumpl['total']),
  '__CUMPL_EB__': fmt(cumpl['en_bins']), '__CUMPL_REQ__': fmt(cumpl['req']), '__CUMPL_FALTAN__': fmt(cumpl['faltan']),
  '__CUMPL_A__': num_es(cumpl['A']), '__CUMPL_B__': num_es(cumpl['B']),
  # Concentración
  '__CONC_N__': str(conc_n), '__CONC_TOP3__': conc_top3, '__CONC_PCT__': str(conc_top3_pct),
  # Corridas
  '__CORR_VIS__': fmt(corr['visits']), '__CORR_CONS__': fmt(corr['cons']),
  '__CORR_PCT__': str(round(corr['pct'])), '__CORR_N__': fmt(corr['corridas']),
  '__CORR_MULTI__': fmt(corr['multi']), '__CORR_PROM__': num_es(corr['prom']),
  '__CORR_DESC__': corr['desc'], '__CORR_MAX__': str(corr['maxv']),
}

interactivo = plantilla
for k, v in valores.items():
    interactivo = interactivo.replace(k, str(v))

# Verificar que no quedaron marcadores sin reemplazar
faltantes = set(re.findall(r'__[A-Z_0-9]+__', interactivo))
if faltantes:
    aviso("[AVISO] Marcadores sin reemplazar: " + ", ".join(sorted(faltantes)))

with open(ruta(SALIDA_INTERACTIVO), 'w', encoding='utf-8') as f:
    f.write(interactivo)
aviso("OK -> " + SALIDA_INTERACTIVO)

# =====================================================================
#  VERSIÓN ESTÁTICA PARA CORREO (sin JavaScript, gráficas en SVG)
# =====================================================================
aviso("Generando versión estática para correo...")
COL = dict(A='#c8412b', B='#d99a2b', C='#3f8f6e', piso='#1d6fb8', altura='#c8412b', slate='#5a6a76')

def svg_vbar(cats, vals, colors, w=440, h=240, pad=38):
    mx = max(vals) if vals else 1; n = len(vals); gap = 18
    bw = (w-2*pad-gap*(n-1))/n; ph = h-pad-24
    o = [f'<svg viewBox="0 0 {w} {h}" xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto">']
    for i in range(5):
        gy = pad/2+ph*i/4; val = mx*(4-i)/4
        o.append(f'<line x1="{pad}" y1="{gy:.0f}" x2="{w-8}" y2="{gy:.0f}" stroke="#eef1f3"/>'
                 f'<text x="{pad-6}" y="{gy+4:.0f}" font-family="Inter,sans-serif" font-size="11" fill="#475560" text-anchor="end">{fmt(round(val))}</text>')
    for i,(c,v,col) in enumerate(zip(cats,vals,colors)):
        x = pad+i*(bw+gap); bh = ph*v/mx if mx else 0; y = pad/2+ph-bh
        o.append(f'<rect x="{x:.0f}" y="{y:.0f}" width="{bw:.0f}" height="{bh:.0f}" rx="6" fill="{col}"/>'
                 f'<text x="{x+bw/2:.0f}" y="{h-6}" font-family="Inter,sans-serif" font-size="12" fill="#475560" text-anchor="middle">{htmlmod.escape(c)}</text>')
    o.append('</svg>'); return ''.join(o)

def svg_donut(vals, colors, labels, w=300, h=240):
    cx, cy, r, rin = 150, 110, 80, 50; total = sum(vals)
    o = [f'<svg viewBox="0 0 {w} {h}" xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto">']; ang = -90
    for v, col in zip(vals, colors):
        frac = v/total if total else 0; a2 = ang+frac*360; la = 1 if frac > 0.5 else 0
        x1 = cx+r*math.cos(math.radians(ang)); y1 = cy+r*math.sin(math.radians(ang))
        x2 = cx+r*math.cos(math.radians(a2)); y2 = cy+r*math.sin(math.radians(a2))
        xi2 = cx+rin*math.cos(math.radians(a2)); yi2 = cy+rin*math.sin(math.radians(a2))
        xi1 = cx+rin*math.cos(math.radians(ang)); yi1 = cy+rin*math.sin(math.radians(ang))
        o.append(f'<path d="M{x1:.1f},{y1:.1f} A{r},{r} 0 {la} 1 {x2:.1f},{y2:.1f} L{xi2:.1f},{yi2:.1f} A{rin},{rin} 0 {la} 0 {xi1:.1f},{yi1:.1f} Z" fill="{col}"/>'); ang = a2
    lx = 30; ly = 212
    for lab, col in zip(labels, colors):
        o.append(f'<rect x="{lx}" y="{ly}" width="11" height="11" rx="2" fill="{col}"/>'
                 f'<text x="{lx+16}" y="{ly+10}" font-family="Inter,sans-serif" font-size="12" fill="#475560">{htmlmod.escape(lab)}</text>'); lx += 130
    o.append('</svg>'); return ''.join(o)

def svg_stack(cats, pv, av, w=440, h=300, pad=44):
    mx = max(p+a for p,a in zip(pv,av)) if cats else 1; n = len(cats); gap = 22
    bw = (w-2*pad-gap*(n-1))/n; ph = h-pad-30
    o = [f'<svg viewBox="0 0 {w} {h}" xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto">']
    for i in range(5):
        gy = pad/2+ph*i/4; val = mx*(4-i)/4
        o.append(f'<line x1="{pad}" y1="{gy:.0f}" x2="{w-8}" y2="{gy:.0f}" stroke="#eef1f3"/>'
                 f'<text x="{pad-6}" y="{gy+4:.0f}" font-family="Inter,sans-serif" font-size="11" fill="#475560" text-anchor="end">{fmt(round(val))}</text>')
    for i, c in enumerate(cats):
        x = pad+i*(bw+gap); p_h = ph*pv[i]/mx; a_h = ph*av[i]/mx; yb = pad/2+ph
        o.append(f'<rect x="{x:.0f}" y="{yb-p_h:.0f}" width="{bw:.0f}" height="{p_h:.0f}" rx="4" fill="{COL["piso"]}"/>'
                 f'<rect x="{x:.0f}" y="{yb-p_h-a_h:.0f}" width="{bw:.0f}" height="{a_h:.0f}" rx="4" fill="{COL["altura"]}"/>'
                 f'<text x="{x+bw/2:.0f}" y="{h-12}" font-family="Inter,sans-serif" font-size="12" fill="#475560" text-anchor="middle">{htmlmod.escape(c)}</text>')
    o.append('</svg>'); return ''.join(o)

estatico = interactivo  # partimos del interactivo ya relleno

def reemplazar(viejo, nuevo):
    global estatico
    if viejo in estatico:
        estatico = estatico.replace(viejo, nuevo, 1)

# Contadores que el JS animaba -> valores fijos
pctA_alt = round(100*a_altura/a_total) if a_total else 0
reemplazar('<div class="sig-num tnum" id="sigNum">0</div>', f'<div class="sig-num tnum" id="sigNum">{fmt(a_altura)}</div>')
reemplazar('<b id="pPctA">0%</b>', f'<b id="pPctA">{pctA_alt}%</b>')
reemplazar('<b id="pSkus">0</b>', f'<b id="pSkus">{fmt(S["a_altura_skus"])}</b>')
reemplazar('<b id="pPedidos">0</b>', f'<b id="pPedidos">{S["a_altura_pedidos"]}</b>')
kpis = [(fmt(total_und),'Unidades pendientes',f"{fmt(len(pend))} líneas · {S['n_pedidos']} pedidos",False),
        (fmt(S['abc_und']['A']),'Son Tipo A',f"{round(100*S['abc_und']['A']/total_und) if total_und else 0}% del pendiente total",True),
        (fmt(piso_und),'A piso (BINS)',f"{round(100*piso_und/total_und) if total_und else 0}% listo para tomar",False),
        (fmt(altura_und),'En altura',f"{round(100*altura_und/total_und) if total_und else 0}% requiere reabasto",False)]
reemplazar('<div class="kpis" id="kpiRow"></div>',
           '<div class="kpis" id="kpiRow">' + ''.join(
               f'<div class="kpi {"accent" if a else ""}"><div class="v tnum">{v}</div><div class="l">{l}</div><div class="d">{d}</div></div>'
               for v,l,d,a in kpis) + '</div>')
reemplazar('<div class="v tnum" id="oUnd">0</div>', f'<div class="v tnum" id="oUnd">{fmt(a_altura)}</div>')
reemplazar('<div class="v tnum" id="oPct">0%</div>', f'<div class="v tnum" id="oPct">{round(100*a_altura/total_und) if total_und else 0}%</div>')
reemplazar('<div class="v tnum" id="oPed">0</div>', f'<div class="v tnum" id="oPed">{S["a_altura_pedidos"]}</div>')
mb = ''
for t in ['A','B','C']:
    r = matriz[t]; tt = r['piso']+r['altura']; pe = round(100*r['piso']/tt) if tt else 0
    mb += (f'<tr><td><span class="chip {t}">{t}</span></td><td>{fmt(r["piso"])}</td><td>{fmt(r["altura"])}</td>'
           f'<td>{fmt(tt)}</td><td><span class="minibar" style="width:{pe*0.55:.0f}px"></span>{pe}%</td></tr>')
mb += (f'<tr><td>TOTAL</td><td>{fmt(piso_und)}</td><td>{fmt(altura_und)}</td><td>{fmt(total_und)}</td>'
       f'<td>{round(100*piso_und/total_und) if total_und else 0}%</td></tr>')
reemplazar('<tbody id="matrixBody"></tbody>', f'<tbody id="matrixBody">{mb}</tbody>')

# OID
gt = sum(o['total'] for o in por_oid); gp = sum(o['piso'] for o in por_oid)
gpct = round(100*gp/gt) if gt else 0
def color_oid(p): return COL['C'] if p>=60 else (COL['piso'] if p>=45 else (COL['B'] if p>=35 else COL['A']))
reemplazar('<div class="og-big tnum" id="oidGlobal">0%</div>', f'<div class="og-big tnum" id="oidGlobal">{gpct}%</div>')
ohtml = ''
for o in por_oid:
    nombre = re.sub(r'^GOO\d+\s*-?\s*','',o['pedido']).strip() or o['pedido']
    col = color_oid(o['pct'])
    pct_s = str(int(o['pct'])) if float(o['pct'])==int(o['pct']) else num_es(o['pct'])
    ohtml += (f'<div class="oidrow"><span class="oid-id">{o["oid"]}</span>'
              f'<span class="oid-name">{htmlmod.escape(nombre)}<small>{htmlmod.escape(o["pedido"].split(" ")[0])}</small></span>'
              f'<span class="oid-track"><span class="oid-fill" style="width:{o["pct"]}%;background:{col}"></span></span>'
              f'<span class="oid-units"><b>{fmt(o["piso"])}</b> / {fmt(o["total"])}</span>'
              f'<span class="oid-pct" style="color:{col}">{pct_s}%</span></div>')
reemplazar('<div class="oidlist" id="oidList"></div>', f'<div class="oidlist" id="oidList">{ohtml}</div>')

# Reabasto + impacto
reemplazar('<div class="rh-num tnum" id="rTotal">0</div>', f'<div class="rh-num tnum" id="rTotal">{fmt(RB["total"])}</div>')
reemplazar('<div class="rc-num tnum" id="rA">0</div>', f'<div class="rc-num tnum" id="rA">{fmt(RB["A"]["faltan"])}</div>')
reemplazar('<div class="rc-num tnum" id="rB">0</div>', f'<div class="rc-num tnum" id="rB">{fmt(RB["B"]["faltan"])}</div>')
reemplazar('<div class="rc-num tnum" id="rC">0</div>', f'<div class="rc-num tnum" id="rC">{fmt(RB["C"]["fuera_rack"])}</div>')
reemplazar('<div class="ba-fill" id="impAfterBar" style="width:0%"></div>',
           f'<div class="ba-fill" id="impAfterBar" style="width:{pct_after}%"></div>')
reemplazar('<div class="iv tnum" id="impToques">0</div>', f'<div class="iv tnum" id="impToques">{fmt(S["a_altura_skus"])}</div>')
reemplazar('<span id="impAltRed">0</span>', f'<span id="impAltRed">{round(100*(altura_und-altura_after)/altura_und) if altura_und else 0}</span>')
reemplazar('<div class="iv tnum" id="impPedidos">0</div>', f'<div class="iv tnum" id="impPedidos">{S["a_altura_pedidos"]}</div>')
reemplazar('<span id="impAcuPct">0</span>', f'<span id="impAcuPct">{pctA_alt}</span>')
reemplazar('<span id="impDestraba">0</span>', f'<span id="impDestraba">{round(100*a_altura/total_und) if total_und else 0}</span>')
reemplazar('<span id="impPisoB">0</span>', f'<span id="impPisoB">{fmt(piso_und)}</span>')
reemplazar('<span id="impPisoA" style="color:var(--piso)">0</span>', f'<span id="impPisoA" style="color:var(--piso)">{fmt(pa)}</span>')

# Gráficas -> SVG
reemplazar('<canvas id="chAbc"></canvas>', svg_vbar(['Tipo A','Tipo B','Tipo C'],
           [S['abc_und']['A'],S['abc_und']['B'],S['abc_und']['C']], [COL['A'],COL['B'],COL['C']]))
reemplazar('<canvas id="chUbic"></canvas>', svg_donut([piso_und,altura_und],[COL['piso'],COL['altura']],
           ['A piso (listo)','En altura']))
reemplazar('<canvas id="chStack"></canvas>', svg_stack(['Tipo A','Tipo B','Tipo C'],
           [matriz['A']['piso'],matriz['B']['piso'],matriz['C']['piso']],
           [matriz['A']['altura'],matriz['B']['altura'],matriz['C']['altura']]))
reemplazar('<canvas id="chReabasto"></canvas>', svg_vbar(['Tipo A','Tipo B','Tipo C'],
           [RB['A']['faltan'],RB['B']['faltan'],RB['C']['fuera_rack']], [COL['A'],COL['B'],COL['C']]))
reemplazar('<canvas id="chImpacto"></canvas>', svg_stack(['Hoy','Con la mejora'],
           [piso_und,pa],[altura_und,altura_after]))

# Quitar todo el JavaScript
estatico = re.sub(r'<script src="https://cdnjs[^"]*"></script>', '', estatico)
estatico = re.sub(r'<script>.*?</script>', '', estatico, flags=re.S)

with open(ruta(SALIDA_CORREO), 'w', encoding='utf-8') as f:
    f.write(estatico)
aviso("OK -> " + SALIDA_CORREO)

# =====================================================================
#  HERRAMIENTA DE REABASTECIMIENTO (lista de tareas: qué bajar, de dónde, a dónde)
# =====================================================================
if os.path.exists(ruta(PLANTILLA_HERRAMIENTA)):
    aviso("Generando herramienta de reabastecimiento...")

    def pasillo_de(ubic):
        u = str(ubic).upper()
        return u.split('-')[0] if '-' in u else u
    def corrida_de(sku):
        s = str(sku); pp = s.rsplit('-', 1); return pp[0] if len(pp) == 2 else s
    def talla_de(sku):
        s = str(sku); pp = s.rsplit('-', 1); return pp[1] if len(pp) == 2 else ''

    inv['corrida'] = inv['Producto'].apply(corrida_de)
    grupos = []
    for tipo in ['A', 'B']:
        sub = inv[inv['Tipo'] == tipo]; needed = req[tipo]
        for corr, gc in sub.groupby('corrida'):
            bins = gc[gc['en_piso']]
            ubic_counts = bins.groupby('Ubicación')['CID'].nunique().sort_values(ascending=False)
            if len(ubic_counts) == 1:
                destino = str(ubic_counts.index[0]).strip(); estado = 'fija'; dispersas = []
            elif len(ubic_counts) > 1:
                destino = str(ubic_counts.index[0]).strip(); estado = 'consolidar'
                dispersas = [str(u).strip() for u in ubic_counts.index]
            else:
                destino = None; estado = 'nueva'; dispersas = []
            # ---- Selección de cajas a bajar ----
            # Primero calculamos cuántas cajas faltan por talla.
            faltan_por_talla = {}
            for prod, gp in gc.groupby('Producto'):
                pbins = gp[gp['en_piso']]['CID'].nunique()
                prack = gp[(~gp['en_piso']) & (gp['Area'] == 'MEZZANINE RACK')]
                falta = max(0, needed - pbins)
                if falta > 0 and len(prack) > 0:
                    faltan_por_talla[prod] = falta

            cajas = []
            # Origen de altura = cajas que NO están en piso y están en RACK
            rack_corr = gc[(~gc['en_piso']) & (gc['Area'] == 'MEZZANINE RACK')]
            # ¿Hay UNA ubicación de altura que tenga suficientes cajas de TODAS las tallas que faltan?
            ubic_completa = None
            for ubic, gu in rack_corr.groupby('Ubicación'):
                if all(gu[gu['Producto'] == prod]['CID'].nunique() >= falta
                       for prod, falta in faltan_por_talla.items()):
                    ubic_completa = ubic
                    break

            for prod, falta in faltan_por_talla.items():
                prack = rack_corr[rack_corr['Producto'] == prod]
                if ubic_completa is not None:
                    # Tomar cajas de la ubicación que tiene la corrida completa (sin importar unidades)
                    pr = prack[prack['Ubicación'] == ubic_completa]
                    rc = pr.groupby('CID').agg(Saldo=('Saldo', 'max'),
                                               Ubic=('Ubicación', 'first')).reset_index()
                    rc = rc.sort_values('Saldo', ascending=False).head(falta)
                else:
                    # Regla normal: las cajas con más unidades disponibles
                    rc = prack.groupby('CID').agg(Saldo=('Saldo', 'max'),
                                                  Ubic=('Ubicación', 'first')).reset_index()
                    rc = rc.sort_values('Saldo', ascending=False).head(falta)
                for _, c in rc.iterrows():
                    cajas.append({'sku': prod, 'talla': talla_de(prod),
                                  'origen': str(c['Ubic']).strip(), 'pasillo': pasillo_de(c['Ubic']),
                                  'cid': str(c['CID']), 'uds': int(c['Saldo'])})
            if not cajas: continue
            cajas.sort(key=lambda x: (x['pasillo'], x['origen'], -x['uds']))
            desc = ''
            if 'Descripción' in pend.columns:
                fila = pend[pend['Producto'] == cajas[0]['sku']]
                if len(fila): desc = str(fila['Descripción'].iloc[0])
            grupos.append({'corrida': corr, 'tipo': tipo, 'desc': desc, 'destino': destino,
                           'estado': estado, 'ubic_dispersas': dispersas, 'n_cajas': len(cajas),
                           'uds': sum(c['uds'] for c in cajas),
                           'completa_en': (str(ubic_completa).strip() if ubic_completa is not None else None),
                           'pasillo_origen': min(c['pasillo'] for c in cajas), 'cajas': cajas})
    orden_tipo = {'A': 0, 'B': 1}
    grupos.sort(key=lambda g: (orden_tipo[g['tipo']], g['pasillo_origen'], g['corrida']))

    # ===== COMPACTAR: cajas dispersas del mismo producto (A: >2 cajas, B: >1 caja) =====
    inv['_pick']  = pd.to_numeric(inv['Pickeado'], errors='coerce').fillna(0)
    inv['_total'] = pd.to_numeric(inv['Total'],    errors='coerce').fillna(0)
    inv['_saldo'] = pd.to_numeric(inv['Saldo'],    errors='coerce').fillna(0)
    UMBRAL_CAJAS = {'A': 2, 'B': 1}   # entra si tiene > umbral cajas EN MEZZANINE BINS
    compactar = []
    for prod, gp in inv.groupby('Producto'):
        tipo = str(gp['Tipo'].mode().iloc[0]).strip()
        if tipo not in UMBRAL_CAJAS:
            continue
        gp_bins = gp[gp['Area'] == 'MEZZANINE BINS']   # solo piso (BINS): ahi se compacta
        if len(gp_bins) <= UMBRAL_CAJAS[tipo]:
            continue
        cajas_c = []
        for _, r in gp_bins.iterrows():
            pick = int(r['_pick'])
            cajas_c.append({'cid': str(r['CID']), 'ubic': str(r['Ubicación']).strip(),
                            'area': str(r['Area']).strip(), 'total': int(r['_total']),
                            'pick': pick, 'saldo': int(r['_saldo']),
                            'accion': 'compactar' if pick > 0 else 'correcta'})
        n_comp = sum(1 for c in cajas_c if c['accion'] == 'compactar')
        if n_comp == 0:
            continue
        cajas_c.sort(key=lambda c: (0 if c['accion'] == 'compactar' else 1, -c['saldo']))
        desc = str(gp['Descrip'].iloc[0]) if 'Descrip' in gp.columns else ''
        compactar.append({'sku': str(prod), 'desc': desc, 'talla': talla_de(prod),
                          'tipo': tipo, 'n_cajas': len(cajas_c), 'n_comp': n_comp,
                          'cajas': cajas_c})
    compactar.sort(key=lambda p: ({'A': 0, 'B': 1}.get(p['tipo'], 2), -p['n_comp'], p['sku']))
    compactar_json = json.dumps(compactar, ensure_ascii=False, separators=(',', ':'))

    with open(ruta(PLANTILLA_HERRAMIENTA), encoding='utf-8') as f:
        tpl_h = f.read()
    grupos_json = json.dumps(grupos, ensure_ascii=False, separators=(',', ':'))
    herramienta = tpl_h.replace('__GRUPOS_REABASTO__', grupos_json)
    herramienta = herramienta.replace('__GRUPOS_COMPACTAR__', compactar_json)
    with open(ruta(SALIDA_HERRAMIENTA), 'w', encoding='utf-8') as f:
        f.write(herramienta)
    n_cajas = sum(g['n_cajas'] for g in grupos)
    n_comp_cajas = sum(p['n_comp'] for p in compactar)
    aviso("OK -> " + SALIDA_HERRAMIENTA + f"  ({len(grupos)} corridas, {n_cajas} cajas"
          f"; compactar: {len(compactar)} productos, {n_comp_cajas} cajas)")

# Limpiar temporal
try: os.remove(ruta('_contexto_tmp.json'))
except OSError: pass

print("="*70)
if primera_vez:
    print("  NOTA: Es la primera corrida, así que el AVANCE aparece en cero")
    print("        (no había corte anterior con qué comparar). La próxima vez")
    print("        que generes con datos nuevos, el avance se calculará solo.")
print("  LISTO. Se generaron los archivos en esta carpeta:")
print("    - " + SALIDA_INTERACTIVO + "   (dashboard para presentar / navegador)")
print("    - " + SALIDA_CORREO + "   (informe para adjuntar al correo)")
print("    - " + SALIDA_HERRAMIENTA + "   (lista de tareas de reabastecimiento)")
print("="*70)
input("\nPresiona ENTER para cerrar...")
