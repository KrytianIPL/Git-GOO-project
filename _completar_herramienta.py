# =====================================================================
#  Auxiliar de generar_dashboard.py
#  Genera SOLO Herramienta_Reabastecimiento_GOO.html usando exactamente
#  el mismo codigo del bloque "HERRAMIENTA" del script principal.
#  Existe porque el script completo puede exceder el limite de tiempo
#  de la ejecucion automatica; aqui se completa ese tercer archivo.
#  NO sustituye a generar_dashboard.py (que genera dashboard e informe).
# =====================================================================
import os, json
import pandas as pd

CARPETA = os.path.dirname(os.path.abspath(__file__))
def ruta(n): return os.path.join(CARPETA, n)
def aviso(m): print("  " + m)

ARCH_INVENTARIO = "Para_analisis_de_GOO_proyecto.xlsx"
ARCH_PENDIENTE  = "Analisis_de_progreso_ABC_GOO.xlsx"
HOJA_PENDIENTE_OPCIONES = ["Pendiente por pick", "Pendiente por Pick", "Pendiente por PICK"]
PLANTILLA_HERRAMIENTA = "plantilla_herramienta.html"
SALIDA_HERRAMIENTA = "Herramienta_Reabastecimiento_GOO.html"

# ---- LECTURA (identica al script principal) ----
aviso("Leyendo inventario...")
inv = pd.read_excel(ruta(ARCH_INVENTARIO))
inv['Area'] = inv['Area'].astype(str).str.strip()
inv['Tipo'] = inv['Tipo (ABC)'].astype(str).str.strip()
inv['CID']  = inv['CID'].astype(str)

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

req = {'A': 2, 'B': 1}

# ===================== BLOQUE HERRAMIENTA (verbatim del script) =====================
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
            faltan_por_talla = {}
            for prod, gp in gc.groupby('Producto'):
                pbins = gp[gp['en_piso']]['CID'].nunique()
                prack = gp[(~gp['en_piso']) & (gp['Area'] == 'MEZZANINE RACK')]
                falta = max(0, needed - pbins)
                if falta > 0 and len(prack) > 0:
                    faltan_por_talla[prod] = falta

            cajas = []
            rack_corr = gc[(~gc['en_piso']) & (gc['Area'] == 'MEZZANINE RACK')]
            ubic_completa = None
            for ubic, gu in rack_corr.groupby('Ubicación'):
                if all(gu[gu['Producto'] == prod]['CID'].nunique() >= falta
                       for prod, falta in faltan_por_talla.items()):
                    ubic_completa = ubic
                    break

            for prod, falta in faltan_por_talla.items():
                prack = rack_corr[rack_corr['Producto'] == prod]
                if ubic_completa is not None:
                    pr = prack[prack['Ubicación'] == ubic_completa]
                    rc = pr.groupby('CID').agg(Saldo=('Saldo', 'max'),
                                               Ubic=('Ubicación', 'first')).reset_index()
                    rc = rc.sort_values('Saldo', ascending=False).head(falta)
                else:
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
else:
    print("[ERROR] No encontre la plantilla de la herramienta.")
