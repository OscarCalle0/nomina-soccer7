"""
app.py — Sistema de Nómina GRANDA VARGAS SAS / Soccer 7
App Streamlit completa con 4 módulos:
  1. Procesar quincena
  2. Colaboradores
  3. Histórico
  4. Nómina electrónica
"""

import streamlit as st
import pandas as pd
import os, io, json, zipfile
from datetime import datetime, date, timedelta
from copy import deepcopy

# ── Configuración de página ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Nómina Soccer 7",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Imports locales ───────────────────────────────────────────────────────────
from datos import (
    cargar_colaboradores, guardar_colaboradores, agregar_colaborador,
    actualizar_colaborador, retirar_colaborador, get_colaborador_por_reloj,
    cargar_historico, guardar_quincena_historico, historico_a_dataframe,
    inicializar_datos, Colaborador, DATA_DIR
)
from colilla_pdf import generar_colilla_pdf, calcular_conceptos_colilla
from nomina_electronica import generar_nomina_electronica_xlsx, preparar_datos_mes_desde_historico

# Motor de nómina (del archivo existente)
import sys
sys.path.insert(0, os.path.dirname(__file__))
from motor_nomina import (
    procesar, calcular,
    crear_reporte_horarios, crear_colilla_pago, crear_resumen_nomina,
)

# ── Inicializar datos ─────────────────────────────────────────────────────────
inicializar_datos()

# ── Estilos CSS — compatibles modo oscuro y claro ───────────────────────────
st.markdown("""
<style>
/* Sidebar */
[data-testid="stSidebar"] { background: #1B3A5C !important; }
[data-testid="stSidebar"] * { color: #e8f0f7 !important; }
[data-testid="stSidebar"] hr { border-color: #2E6DA4 !important; }

/* Títulos */
h1, h2 { color: var(--text-color) !important; }

/* Días del preinforme — colores semáforo legibles en oscuro y claro */
.dia-ok    { border-left: 3px solid #27ae60; padding: 5px 10px; border-radius: 4px;
             background: rgba(39,174,96,0.12); margin: 3px 0; }
.dia-alerta{ border-left: 3px solid #e67e22; padding: 5px 10px; border-radius: 4px;
             background: rgba(230,126,34,0.12); margin: 3px 0; }
.dia-sin   { border-left: 3px solid #95a5a6; padding: 5px 10px; border-radius: 4px;
             background: rgba(149,165,166,0.10); margin: 3px 0; }

/* Texto de horas — alto contraste en ambos modos */
.hora-txt  { font-family: monospace; font-size: 13px;
             color: var(--text-color); font-weight: 500; }
.hora-noct { font-family: monospace; font-size: 12px; color: #5dade2; }
.alerta-txt{ font-family: monospace; font-size: 13px; color: #e67e22; font-weight: 600; }
.ok-txt    { font-size: 12px; color: #27ae60; }

/* Responsive — en móvil ocultar columnas extras */
@media (max-width: 768px) {
    [data-testid="stSidebar"] { display: none; }
    .block-container { padding: 0.5rem !important; }
}
</style>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
    if os.path.exists(logo_path):
        st.image(logo_path, width=110)
    st.markdown("## Sistema de Nómina")
    st.markdown("GRANDA VARGAS SAS")
    st.divider()

    modulo = st.radio("Módulo", [
        "⚙️  Procesar quincena",
        "👥  Colaboradores",
        "📋  Histórico",
        "📊  Nómina electrónica",
    ], label_visibility="collapsed")
    st.divider()
    st.markdown("Soccer 7 · v2.0")

# ═══════════════════════════════════════════════════════════════════════════════
# MÓDULO 1 — PROCESAR QUINCENA
# Flujo simplificado: Período → Archivo+Preinforme → Calcular+Descargar
# ═══════════════════════════════════════════════════════════════════════════════
if "⚙️" in modulo:
    st.title("⚙️ Procesar quincena")
    colaboradores_db = cargar_colaboradores()

    TIPOS_NOV_PI = {
        "— Sin novedad —":              None,
        "Vacaciones":                   "VACAC",
        "Incapacidad EPS":              "INC_EPS",
        "Incapacidad ARL (100%)":       "INC_ARL",
        "Licencia remunerada":          "LIC_REM",
        "Licencia no remunerada":       "LIC_NREM",
        "Día de la familia":            "DIA_FAM",
        "Día compensatorio":            "COMPENS",
        "Calamidad doméstica":          "CALAM",
        "Ausencia injustificada":       "AUS_INJ",
        "Suspensión disciplinaria":     "SUSPEND",
        "Maternidad / Paternidad":      "MAT_PAT",
        "Renuncia / Retiro":            "RENUNCIA",
        "Ingreso nuevo en período":     "INGRESO",
    }
    PCT_DEFAULT = {"INC_EPS":66.66,"LIC_NREM":0.0,"SUSPEND":0.0,"AUS_INJ":0.0,"COMPENS":0.0,"RENUNCIA":100.0,"INGRESO":100.0}
    DIAS_SEM = ["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"]

    # ── PASO 1: Período ──────────────────────────────────────────────────────
    with st.expander("📅 Paso 1 — Período de liquidación", expanded=True):
        c1, c2, c3 = st.columns([1,1,2])
        fecha_ini = c1.date_input("Desde", value=date.today().replace(day=1))
        fecha_fin = c2.date_input("Hasta", value=date.today().replace(day=15))
        q_num = "primera" if fecha_ini.day == 1 else "segunda"
        c3.info(f"**{q_num.capitalize()} quincena** de {fecha_ini.strftime('%B %Y').capitalize()} · Meta: 88 horas")

    # ── PASO 2: Archivo + Preinforme editable ────────────────────────────────
    with st.expander("📂 Paso 2 — Archivo del reloj y revisión de marcaciones", expanded=True):
        archivo = st.file_uploader(
            "Sube el reporte del Zkteco K50",
            type=["xls","xlsx","csv","txt"],
            help="Acepta .xls, .xlsx y .csv del reloj biométrico"
        )

        if archivo:
            ext = archivo.name.lower().split(".")[-1]
            df_reloj = None
            try:
                if ext == "xls":       df_reloj = pd.read_excel(archivo, engine="xlrd")
                elif ext in ["xlsx","xlsm"]: df_reloj = pd.read_excel(archivo, engine="openpyxl")
                else:
                    for sep in [",",";"," \t"]:
                        try:
                            archivo.seek(0)
                            tmp = pd.read_csv(archivo, sep=sep, encoding="utf-8-sig")
                            if len(tmp.columns) >= 4: df_reloj = tmp; break
                        except: pass

                if df_reloj is None:
                    st.error("No se pudo leer el archivo."); st.stop()

                df_reloj.columns = [str(c).strip() for c in df_reloj.columns]
                if "Numero" in df_reloj.columns and "Número" not in df_reloj.columns:
                    df_reloj.rename(columns={"Numero":"Número"}, inplace=True)

                p_ini_dt = datetime.combine(fecha_ini, datetime.min.time())
                p_fin_dt = datetime.combine(fecha_fin, datetime.max.time())

                # Solo reprocesar si es un archivo nuevo
                archivo_nombre = archivo.name + str(fecha_ini) + str(fecha_fin)
                if st.session_state.get("archivo_cargado") != archivo_nombre:
                    resultados_raw = procesar(df_reloj, p_ini_dt, p_fin_dt)
                    st.session_state.update({
                        "df_reloj": df_reloj,
                        "resultados_raw": resultados_raw,
                        "p_ini_dt": p_ini_dt,
                        "p_fin_dt": p_fin_dt,
                        "novedades": {e["nombre"]: [] for e in resultados_raw},
                        "novedades_pi": {},
                        "correcciones_pendientes": {},
                        "archivo_cargado": archivo_nombre,
                    })
                    st.success(f"✅ **{archivo.name}** · {len(df_reloj)} marcaciones · {len(resultados_raw)} colaboradoras")
                else:
                    st.success(f"✅ **{archivo.name}** ya cargado — ediciones preservadas")

            except Exception as e:
                st.error(f"Error leyendo el archivo: {e}"); st.stop()

        # ── Preinforme editable ──────────────────────────────────────────────
        resultados_raw = st.session_state.get("resultados_raw", [])
        if not resultados_raw:
            st.info("Sube el archivo del reloj para ver el preinforme.")
        else:
            p_ini_dt = st.session_state["p_ini_dt"]
            p_fin_dt = st.session_state["p_fin_dt"]

            # Métricas globales
            total_alertas = sum(len([d for d in e["dias"] if d["tiene"] and (d["ef"]=="missing" or d["sf"]=="missing")]) for e in resultados_raw)
            total_sin     = sum(len([d for d in e["dias"] if not d["tiene"]]) for e in resultados_raw)
            total_horas   = sum(sum(d["trab"]*24 for d in e["dias"]) for e in resultados_raw)

            m1,m2,m3,m4 = st.columns(4)
            m1.metric("Colaboradoras", len(resultados_raw))
            m2.metric("Horas totales", f"{total_horas:.0f}h")
            m3.metric("Alertas ⚠️", total_alertas, delta="requieren revisión" if total_alertas else None, delta_color="inverse")
            m4.metric("Sin registro", total_sin)
            st.divider()

            novedades_pi = st.session_state.get("novedades_pi", {})

            for emp_idx, emp in enumerate(resultados_raw):
                col_db    = get_colaborador_por_reloj(emp["nombre"], colaboradores_db)
                horas_tot = sum(d["trab"]*24 for d in emp["dias"])
                n_alertas = len([d for d in emp["dias"] if d["tiene"] and (d["ef"]=="missing" or d["sf"]=="missing")])
                n_sin     = len([d for d in emp["dias"] if not d["tiene"]])
                n_turnos  = sum(len(d.get("turnos",[])) for d in emp["dias"])
                diff_h    = horas_tot - 88

                icono  = "⚠️" if n_alertas else "✅"
                bd_txt = "" if col_db else " · ❌ No en BD"
                titulo = f"{icono} **{emp['nombre']}**  {horas_tot:.1f}h ({diff_h:+.1f}h) · {n_turnos} turnos · {n_alertas} alertas · {n_sin} sin reg{bd_txt}"

                with st.expander(titulo, expanded=bool(n_alertas)):
                    if not col_db:
                        st.warning(f"⚠️ {emp['nombre']} no está en la base de datos. Agrégala en el módulo Colaboradores.")

                    for d in emp["dias"]:
                        dia_nom   = DIAS_SEM[d["fecha"].weekday()]
                        fecha_lbl = f"{dia_nom} {d['fecha'].day:02d}/{d['fecha'].month:02d}"
                        dk        = d["dk"]
                        nov_key   = f"{emp['nombre']}||{dk}"
                        horas_dia = d["trab"]*24
                        noct_dia  = d["noct"]*24
                        turnos    = d.get("turnos", [])

                        # Barra de estado del día
                        if not d["tiene"]:
                            st.markdown(f'<div class="dia-sin"><span style="font-weight:500">{fecha_lbl}</span> &nbsp; <span style="opacity:.7">⚪ Sin registro</span></div>', unsafe_allow_html=True)
                        elif d["ef"]=="missing" or d["sf"]=="missing":
                            st.markdown(f'<div class="dia-alerta"><span style="font-weight:500">{fecha_lbl}</span> &nbsp; <span class="alerta-txt">⚠️ {horas_dia:.2f}h · {len(turnos)} turno(s)</span></div>', unsafe_allow_html=True)
                        else:
                            noct_str = f' · <span class="hora-noct">🌙 {noct_dia:.1f}h noct</span>' if noct_dia > 0 else ""
                            st.markdown(f'<div class="dia-ok"><span style="font-weight:500">{fecha_lbl}</span> &nbsp; <span class="hora-txt">🟢 {horas_dia:.2f}h · {len(turnos)} turno(s)</span>{noct_str}</div>', unsafe_allow_html=True)

                        # Turnos del día
                        for t_idx, turno in enumerate(turnos):
                            t_trab = 0.0
                            if turno["entrada"] and turno["salida"]:
                                t_trab = (turno["salida"]-turno["entrada"]).total_seconds()/3600
                            e_str = turno["entrada"].strftime("%H:%M") if turno["entrada"] else "❌"
                            s_str = turno["salida"].strftime("%H:%M")  if turno["salida"]  else "❌"

                            tc1,tc2,tc3,tc4,tc5,tc6,tc7 = st.columns([0.7,1,1,1.3,1.3,0.8,0.4])
                            tc1.markdown(f'<span style="font-size:11px;opacity:.7">T{t_idx+1}</span>', unsafe_allow_html=True)
                            e_cls = "alerta-txt" if turno["alerta"]=="sin_entrada" else "hora-txt"
                            s_cls = "alerta-txt" if turno["alerta"]=="sin_salida"  else "hora-txt"
                            tc2.markdown(f'<span class="{e_cls}">E: {e_str}</span>', unsafe_allow_html=True)
                            tc3.markdown(f'<span class="{s_cls}">S: {s_str}</span>', unsafe_allow_html=True)

                            nueva_e = tc4.text_input(" ", placeholder="Nueva E HH:MM", key=f"ne_{emp_idx}_{dk}_{t_idx}", label_visibility="collapsed")
                            nueva_s = tc5.text_input(" ", placeholder="Nueva S HH:MM", key=f"ns_{emp_idx}_{dk}_{t_idx}", label_visibility="collapsed")
                            if t_trab > 0:
                                tc6.markdown(f'<span class="hora-txt">{t_trab:.1f}h</span>', unsafe_allow_html=True)

                            if tc7.button("🗑️", key=f"del_{emp_idx}_{dk}_{t_idx}", help="Eliminar turno"):
                                df_act = st.session_state["df_reloj"]
                                to_drop = set()
                                if turno["entrada"]: to_drop.add(turno["entrada"].strftime("%d/%m/%Y %H:%M:%S"))
                                if turno["salida"]:  to_drop.add(turno["salida"].strftime("%d/%m/%Y %H:%M:%S"))
                                if to_drop:
                                    mask = ~((df_act["Nombre"]==emp["nombre"]) & (df_act["Tiempo"].astype(str).isin(to_drop)))
                                    df_filt = df_act[mask].reset_index(drop=True)
                                    nuevos_res = procesar(df_filt, p_ini_dt, p_fin_dt)
                                    st.session_state["df_reloj"]      = df_filt
                                    st.session_state["resultados_raw"] = nuevos_res
                                    st.success("✅ Turno eliminado")
                                    st.rerun()

                            # Guardar correcciones pendientes
                            if nueva_e.strip() or nueva_s.strip():
                                corr = st.session_state.setdefault("correcciones_pendientes", {})
                                corr[f"{emp_idx}_{dk}_{t_idx}"] = {
                                    "nombre": emp["nombre"], "emp_id": emp["id"],
                                    "fecha_dt": d["fecha"],
                                    "entrada": nueva_e.strip() or None,
                                    "salida":  nueva_s.strip() or None,
                                }

                        # Agregar turno nuevo
                        a1,a2,a3,_ = st.columns([0.7,1.3,1.3,2])
                        a1.markdown('<span style="font-size:11px;color:#5dade2">➕ nuevo</span>', unsafe_allow_html=True)
                        add_e = a2.text_input(" ", placeholder="Entrada HH:MM", key=f"ae_{emp_idx}_{dk}", label_visibility="collapsed")
                        add_s = a3.text_input(" ", placeholder="Salida HH:MM",  key=f"as_{emp_idx}_{dk}", label_visibility="collapsed")
                        if add_e.strip() or add_s.strip():
                            corr = st.session_state.setdefault("correcciones_pendientes", {})
                            corr[f"add_{emp_idx}_{dk}"] = {
                                "nombre": emp["nombre"], "emp_id": emp["id"],
                                "fecha_dt": d["fecha"],
                                "entrada": add_e.strip() or None,
                                "salida":  add_s.strip() or None,
                            }

                        # Novedad del día (dropdown)
                        nov_actual = novedades_pi.get(nov_key, "— Sin novedad —")
                        opciones_nov = list(TIPOS_NOV_PI.keys())
                        idx_nov = opciones_nov.index(nov_actual) if nov_actual in opciones_nov else 0
                        nov_cols = st.columns([0.7, 3, 1.5])
                        nov_cols[0].markdown('<span style="font-size:11px;opacity:.7">📌 nov</span>', unsafe_allow_html=True)
                        nov_sel = nov_cols[1].selectbox(
                            "", opciones_nov, index=idx_nov,
                            key=f"nov_{emp_idx}_{dk}", label_visibility="collapsed"
                        )
                        # Días de la novedad
                        nov_dias = nov_cols[2].number_input(
                            "", min_value=0.5, max_value=15.0, value=1.0, step=0.5,
                            key=f"novd_{emp_idx}_{dk}", label_visibility="collapsed"
                        ) if nov_sel != "— Sin novedad —" else 1.0

                        if nov_sel != "— Sin novedad —":
                            novedades_pi[nov_key] = {"tipo_desc": nov_sel, "dias": nov_dias}
                        elif nov_key in novedades_pi:
                            del novedades_pi[nov_key]

                        st.markdown("<hr style='margin:4px 0;opacity:.2'>", unsafe_allow_html=True)

                    st.session_state["novedades_pi"] = novedades_pi

                    # Botón aplicar correcciones de esta persona
                    mis_correcciones = {k:v for k,v in st.session_state.get("correcciones_pendientes",{}).items() if v.get("nombre")==emp["nombre"]}
                    if mis_correcciones:
                        st.info(f"📝 {len(mis_correcciones)} corrección(es) pendiente(s)")
                        if st.button(f"✅ Aplicar correcciones de {emp['nombre'].split()[0]}", key=f"apply_{emp_idx}", type="primary"):
                            df_act = st.session_state["df_reloj"]
                            nuevas_filas = []
                            errs = []
                            for ckey, corr in mis_correcciones.items():
                                fd = corr["fecha_dt"]
                                if corr["entrada"]:
                                    try:
                                        hh,mm = map(int, corr["entrada"].split(":"))
                                        ts = fd.replace(hour=hh, minute=mm, second=0)
                                        nuevas_filas.append({"Número":corr["emp_id"],"Nombre":corr["nombre"],"Tiempo":ts.strftime("%d/%m/%Y %H:%M:%S"),"Estado":"Entrada","Dispositivos":"MANUAL","Tipo de Registro":0})
                                    except: errs.append(f"Entrada inválida: {corr['entrada']}")
                                if corr["salida"]:
                                    try:
                                        hh,mm = map(int, corr["salida"].split(":"))
                                        if hh<6 and corr["entrada"] and int(corr["entrada"].split(":")[0])>=12:
                                            ts_s = (fd+timedelta(days=1)).replace(hour=hh,minute=mm,second=0)
                                        else:
                                            ts_s = fd.replace(hour=hh,minute=mm,second=0)
                                        nuevas_filas.append({"Número":corr["emp_id"],"Nombre":corr["nombre"],"Tiempo":ts_s.strftime("%d/%m/%Y %H:%M:%S"),"Estado":"Salida","Dispositivos":"MANUAL","Tipo de Registro":0})
                                    except: errs.append(f"Salida inválida: {corr['salida']}")
                            if errs:
                                for e in errs: st.error(e)
                            else:
                                if nuevas_filas:
                                    df_act = pd.concat([df_act, pd.DataFrame(nuevas_filas)], ignore_index=True)
                                nuevos_res = procesar(df_act, p_ini_dt, p_fin_dt)
                                st.session_state["df_reloj"]      = df_act
                                st.session_state["resultados_raw"] = nuevos_res
                                for k in mis_correcciones: del st.session_state["correcciones_pendientes"][k]
                                st.success("✅ Correcciones aplicadas")
                                st.rerun()

                    # Métricas del colaborador
                    rc1,rc2,rc3,rc4 = st.columns(4)
                    rc1.metric("Horas totales", f"{horas_tot:.1f}h", delta=f"{diff_h:+.1f}h vs 88h", delta_color="normal" if diff_h>=0 else "inverse")
                    rc2.metric("Turnos", n_turnos)
                    rc3.metric("Alertas", n_alertas)
                    rc4.metric("Sin registro", n_sin)

            # Consolidar novedades del preinforme al dict de novedades
            novedades_dict = {e["nombre"]: [] for e in resultados_raw}
            for nov_key, nov_data in novedades_pi.items():
                nombre = nov_key.split("||")[0]
                tipo_codigo = TIPOS_NOV_PI.get(nov_data.get("tipo_desc",""), None)
                if tipo_codigo and nombre in novedades_dict:
                    dias = nov_data.get("dias", 1.0)
                    pct  = PCT_DEFAULT.get(tipo_codigo, 100.0)
                    # Evitar duplicados del mismo tipo
                    ya = any(n["tipo"]==tipo_codigo for n in novedades_dict[nombre])
                    if not ya:
                        novedades_dict[nombre].append({"tipo":tipo_codigo,"dias":dias,"pct":pct,"valor_override":None})
            st.session_state["novedades"] = novedades_dict

    # ── PASO 3: Calcular y descargar ─────────────────────────────────────────
    with st.expander("🚀 Paso 3 — Calcular y descargar", expanded=True):
        if "resultados_raw" not in st.session_state:
            st.info("Primero carga el archivo del reloj en el Paso 2.")
        else:
            # Mostrar resumen de novedades antes de calcular
            novedades_dict = st.session_state.get("novedades", {})
            novedades_activas = [(n,vs) for n,vs in novedades_dict.items() if vs]
            if novedades_activas:
                st.markdown("**Novedades registradas en el preinforme:**")
                for nombre, novs in novedades_activas:
                    for nov in novs:
                        st.caption(f"  📌 {nombre} — {nov['tipo']} · {nov['dias']} días · {nov['pct']:.1f}%")
                st.divider()

            if st.button("🔄 Calcular nómina", type="primary", use_container_width=True):
                p_ini_dt = st.session_state["p_ini_dt"]
                p_fin_dt = st.session_state["p_fin_dt"]
                resultados_raw = st.session_state["resultados_raw"]
                novedades_dict = st.session_state.get("novedades", {})

                resultados_t = []
                for emp in resultados_raw:
                    col_db = get_colaborador_por_reloj(emp["nombre"], colaboradores_db)
                    sal  = (col_db.salario_mensual if col_db.tipo=="empleado" else col_db.valor_hora_prestador) if col_db else 1423500
                    tipo = col_db.tipo if col_db else "empleado"
                    novs = novedades_dict.get(emp["nombre"], [])
                    t    = calcular(emp, sal, tipo, novs)
                    resultados_t.append((emp, t))

                st.session_state["resultados_t"] = resultados_t
                st.success("✅ Nómina calculada")

            if "resultados_t" in st.session_state:
                resultados_t = st.session_state["resultados_t"]
                p_ini_dt = st.session_state["p_ini_dt"]
                p_fin_dt = st.session_state["p_fin_dt"]

                # Tabla resumen
                st.subheader("Resumen")
                rows_res = []
                gran_neto = 0
                for emp, t in resultados_t:
                    col_db = get_colaborador_por_reloj(emp["nombre"], colaboradores_db)
                    if col_db and col_db.tipo=="prestador":
                        neto = t["val_total_prest"]
                    else:
                        sal   = col_db.salario_mensual if col_db else 1423500
                        sal_q = sal/2
                        dias  = t["dias_trab"]
                        aux   = (249095/2)*dias/15
                        ibc   = sal_q+t["val_en"]+t["val_noct"]
                        ded   = ibc*0.08+t["nov_deduccion"]
                        neto  = sal_q+aux+t["val_en"]+t["val_noct"]+t["nov_devengado"]-ded
                    gran_neto += neto
                    estado = f"✅ +{t.get('en_h',0):.1f}h" if t.get("en_h",0)>0 else (f"⚠️ debe {t.get('deu_h',0):.1f}h" if t.get("deu_h",0)>0.1 else "✓ OK")
                    novs_txt = ", ".join([nd.get("desc","").split(" —")[0] for nd in t.get("nov_detalle",[])]) or "—"
                    rows_res.append({"Colaboradora":emp["nombre"],"Horas":f"{t['tot_h']:.1f}h","Extras":f"{t.get('en_h',0):.1f}h","Recargo Noct.":f"${t['val_noct']:,.0f}","Novedades":novs_txt,"Estado":estado,"Neto":f"${neto:,.0f}"})

                st.dataframe(pd.DataFrame(rows_res), hide_index=True, use_container_width=True)
                m1,m2,m3,m4 = st.columns(4)
                m1.metric("Total neto", f"${gran_neto:,.0f}")
                m2.metric("Extras efectivo", f"${sum(t.get('val_ee',0) for _,t in resultados_t):,.0f}")
                m3.metric("Recargo nocturno", f"${sum(t.get('val_noct',0) for _,t in resultados_t):,.0f}")
                m4.metric("Colaboradoras", len(resultados_t))

                st.divider()
                st.subheader("Descargar archivos")
                tag = f"{p_ini_dt.strftime('%Y%m%d')}_{p_fin_dt.strftime('%Y%m%d')}"
                c1,c2,c3,c4 = st.columns(4)

                # Colillas PDF
                with c1:
                    if st.button("📄 Generar colillas PDF", use_container_width=True):
                        lista_c = []
                        for emp,t in resultados_t:
                            col_db = get_colaborador_por_reloj(emp["nombre"], colaboradores_db)
                            cd = {"id":col_db.id,"nombre_completo":col_db.nombre_completo,"cargo":col_db.cargo,"salario_mensual":col_db.salario_mensual,"tipo":col_db.tipo,"banco":col_db.banco,"cuenta":col_db.cuenta,"eps":col_db.eps,"valor_hora_prestador":col_db.valor_hora_prestador} if col_db else {"id":emp["id"],"nombre_completo":emp["nombre"],"cargo":"","salario_mensual":t["salario"],"tipo":t["tipo"],"banco":"","cuenta":"","eps":"","valor_hora_prestador":0}
                            lista_c.append(calcular_conceptos_colilla(t, cd))
                        pdf_bytes = generar_colilla_pdf(lista_c, p_ini_dt, p_fin_dt)
                        st.session_state["pdf_bytes"] = pdf_bytes
                    if st.session_state.get("pdf_bytes"):
                        st.download_button("⬇️ Descargar PDF", data=st.session_state["pdf_bytes"], file_name=f"COLILLAS_{tag}.pdf", mime="application/pdf", use_container_width=True)

                # Reporte horarios
                with c2:
                    crear_reporte_horarios(resultados_t, p_ini_dt, p_fin_dt, f"/tmp/rh_{tag}.xlsx")
                    with open(f"/tmp/rh_{tag}.xlsx","rb") as f_: rh = f_.read()
                    st.download_button("⬇️ Reporte horarios", data=rh, file_name=f"REPORTE_HORARIOS_{tag}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

                # Resumen nómina
                with c3:
                    vd_q = 1750905/2; vd_aux = 249095/2; valentina_neto = vd_q+vd_aux-vd_q*0.08
                    crear_resumen_nomina(resultados_t, valentina_neto, p_ini_dt, p_fin_dt, f"/tmp/rn_{tag}.xlsx")
                    with open(f"/tmp/rn_{tag}.xlsx","rb") as f_: rn = f_.read()
                    st.download_button("⬇️ Resumen nómina", data=rn, file_name=f"RESUMEN_NOMINA_{tag}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

                # ZIP todo
                with c4:
                    if st.button("📦 Todo en ZIP", use_container_width=True):
                        zb = io.BytesIO()
                        with zipfile.ZipFile(zb,"w") as zf:
                            if st.session_state.get("pdf_bytes"): zf.writestr(f"COLILLAS_{tag}.pdf", st.session_state["pdf_bytes"])
                            zf.write(f"/tmp/rh_{tag}.xlsx", f"REPORTE_HORARIOS_{tag}.xlsx")
                            zf.write(f"/tmp/rn_{tag}.xlsx", f"RESUMEN_NOMINA_{tag}.xlsx")
                        st.download_button("⬇️ Descargar ZIP", data=zb.getvalue(), file_name=f"NOMINA_{tag}.zip", mime="application/zip", use_container_width=True)

                # Guardar en histórico
                st.divider()
                if st.button("💾 Guardar quincena en el histórico", type="primary"):
                    periodo_id = f"{p_ini_dt.strftime('%Y-%m-%d')}_{p_fin_dt.strftime('%Y-%m-%d')}"
                    registro = {
                        "id": periodo_id,
                        "periodo_ini": p_ini_dt.strftime("%Y-%m-%d"),
                        "periodo_fin": p_fin_dt.strftime("%Y-%m-%d"),
                        "fecha_procesado": datetime.now().isoformat(),
                        "total_nomina": gran_neto,
                        "total_extras_ef": sum(t.get("val_ee",0) for _,t in resultados_t),
                        "total_recargo": sum(t.get("val_noct",0) for _,t in resultados_t),
                        "colaboradores": [
                            {"nombre":emp["nombre"],"tot_h":t.get("tot_h",0),"noct_h":t.get("noct_h",0),
                             "ext_h":t.get("en_h",0)+t.get("ee_h",0),"en_h":t.get("en_h",0),"ee_h":t.get("ee_h",0),
                             "val_en":t.get("val_en",0),"val_ee":t.get("val_ee",0),"val_noct":t.get("val_noct",0),
                             "val_deu":t.get("val_deu",0),"deu_h":t.get("deu_h",0),"dias_trab":t.get("dias_trab",0),
                             "salario":t.get("salario",0),"tipo":t.get("tipo","empleado"),
                             "val_total_prest":t.get("val_total_prest",0),
                             "nov_devengado":t.get("nov_devengado",0),"nov_deduccion":t.get("nov_deduccion",0),
                             "novedades_desc":" / ".join([nd.get("desc","") for nd in t.get("nov_detalle",[])])}
                            for emp,t in resultados_t
                        ],
                    }
                    guardar_quincena_historico(registro)
                    st.success(f"✅ Quincena {periodo_id} guardada")


# ═══════════════════════════════════════════════════════════════════════════════
# MÓDULO 2 — COLABORADORES
# ═══════════════════════════════════════════════════════════════════════════════
elif "👥" in modulo:
    st.title("👥 Colaboradores")
    colaboradores_db = cargar_colaboradores()

    tab_lista, tab_nuevo, tab_editar = st.tabs(["Lista del personal", "Agregar nuevo", "Editar / Retirar"])

    with tab_lista:
        st.subheader(f"Personal registrado ({len(colaboradores_db)} personas)")
        rows = []
        for c in colaboradores_db:
            rows.append({
                "Cédula": c.id,
                "Nombre": c.nombre_completo,
                "Nombre en reloj": c.nombre_reloj,
                "Cargo": c.cargo,
                "Salario mensual": f"${c.salario_mensual:,.0f}",
                "Tipo": c.tipo.capitalize(),
                "Fecha ingreso": c.fecha_ingreso,
                "Fecha retiro": c.fecha_retiro or "—",
                "Estado": "🟢 Activo" if c.activo else "🔴 Retirado",
                "Banco": c.banco,
                "Cuenta": c.cuenta,
                "EPS": c.eps,
            })
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

    with tab_nuevo:
        st.subheader("Agregar nueva colaboradora")
        with st.form("form_nuevo_col"):
            c1, c2 = st.columns(2)
            with c1:
                ced     = st.text_input("Cédula *")
                nombre  = st.text_input("Nombre completo * (como en colilla)")
                n_reloj = st.text_input("Nombre en el reloj * (exacto)")
                cargo   = st.text_input("Cargo")
                ingreso = st.date_input("Fecha de ingreso", value=date.today())
            with c2:
                tipo     = st.selectbox("Tipo", ["empleado", "prestador"])
                salario  = st.number_input("Salario mensual (COP)", min_value=0, value=1750905, step=10000)
                vh_prest = st.number_input("Valor hora (solo prestador)", min_value=0, value=10000)
                banco    = st.text_input("Banco", value="AHORROS DAVIVIENDA")
                cuenta   = st.text_input("Número de cuenta")
                eps      = st.text_input("EPS", value="SAVIA SALUD")

            notas = st.text_area("Notas", height=60)

            if st.form_submit_button("✅ Agregar colaboradora", type="primary"):
                if not ced or not nombre or not n_reloj:
                    st.error("Cédula, nombre completo y nombre en reloj son obligatorios.")
                else:
                    nuevo = Colaborador(
                        id=ced.strip(), nombre_completo=nombre.strip(),
                        nombre_reloj=n_reloj.strip(), cargo=cargo.strip(),
                        salario_mensual=salario, fecha_ingreso=ingreso.isoformat(),
                        fecha_retiro=None, banco=banco.strip(), cuenta=cuenta.strip(),
                        eps=eps.strip(), tipo=tipo,
                        valor_hora_prestador=vh_prest if tipo=="prestador" else 0,
                        activo=True, notas=notas
                    )
                    if agregar_colaborador(nuevo):
                        st.success(f"✅ {nombre} agregada correctamente")
                        st.rerun()
                    else:
                        st.error(f"Ya existe una colaboradora con cédula {ced}")

    with tab_editar:
        st.subheader("Editar datos o registrar retiro")
        nombres_todos = [f"{c.nombre_completo} ({c.id})" for c in colaboradores_db]
        seleccion = st.selectbox("Selecciona colaboradora", nombres_todos)

        if seleccion:
            idx = nombres_todos.index(seleccion)
            col_sel = colaboradores_db[idx]

            with st.form("form_editar"):
                c1, c2 = st.columns(2)
                with c1:
                    nuevo_nombre  = st.text_input("Nombre completo", value=col_sel.nombre_completo)
                    nuevo_reloj   = st.text_input("Nombre en reloj", value=col_sel.nombre_reloj)
                    nuevo_cargo   = st.text_input("Cargo", value=col_sel.cargo)
                    nuevo_ingreso = st.date_input("Fecha ingreso",
                                                   value=date.fromisoformat(col_sel.fecha_ingreso))
                with c2:
                    nuevo_salario = st.number_input("Salario mensual", value=int(col_sel.salario_mensual), step=10000)
                    nuevo_banco   = st.text_input("Banco", value=col_sel.banco)
                    nuevo_cuenta  = st.text_input("Cuenta", value=col_sel.cuenta)
                    nuevo_eps     = st.text_input("EPS", value=col_sel.eps)
                    nuevo_vh      = st.number_input("Valor hora prestador", value=int(col_sel.valor_hora_prestador))
                nuevas_notas = st.text_area("Notas", value=col_sel.notas)

                st.divider()
                st.markdown("**Registrar retiro** (deja vacío si sigue activa)")
                fecha_retiro_inp = st.date_input("Fecha de retiro",
                                                  value=date.fromisoformat(col_sel.fecha_retiro) if col_sel.fecha_retiro else None)

                c_save, c_retire = st.columns(2)
                with c_save:
                    if st.form_submit_button("💾 Guardar cambios", type="primary"):
                        col_sel.nombre_completo   = nuevo_nombre
                        col_sel.nombre_reloj      = nuevo_reloj
                        col_sel.cargo             = nuevo_cargo
                        col_sel.salario_mensual   = nuevo_salario
                        col_sel.fecha_ingreso     = nuevo_ingreso.isoformat()
                        col_sel.banco             = nuevo_banco
                        col_sel.cuenta            = nuevo_cuenta
                        col_sel.eps               = nuevo_eps
                        col_sel.valor_hora_prestador = nuevo_vh
                        col_sel.notas             = nuevas_notas
                        actualizar_colaborador(col_sel)
                        st.success("✅ Datos actualizados")
                        st.rerun()
                with c_retire:
                    if st.form_submit_button("🔴 Registrar retiro", type="secondary"):
                        if fecha_retiro_inp:
                            retirar_colaborador(col_sel.id, fecha_retiro_inp.isoformat())
                            st.success(f"✅ {col_sel.nombre_completo} retirada el {fecha_retiro_inp}")
                            st.rerun()
                        else:
                            st.error("Selecciona la fecha de retiro")


# ═══════════════════════════════════════════════════════════════════════════════
# MÓDULO 3 — HISTÓRICO
# ═══════════════════════════════════════════════════════════════════════════════
elif "📋" in modulo:
    st.title("📋 Histórico de quincenas")
    historico = cargar_historico()

    if not historico:
        st.info("Aún no hay quincenas guardadas. Procesa la primera quincena en el módulo de nómina.")
    else:
        # Resumen general
        col1, col2, col3 = st.columns(3)
        col1.metric("Quincenas registradas", len(historico))
        total_hist = sum(h.get("total_nomina", 0) for h in historico)
        col2.metric("Total nómina histórica", f"${total_hist:,.0f}")
        ultima = historico[0]
        col3.metric("Última quincena", ultima.get("id", "—").replace("_", " al "))

        st.divider()

        # Lista de quincenas
        st.subheader("Quincenas procesadas")
        df_hist_res = pd.DataFrame([{
            "Período": h.get("id","").replace("_"," al "),
            "Fecha procesado": h.get("fecha_procesado","")[:10],
            "Total nómina": f"${h.get('total_nomina',0):,.0f}",
            "Extras efectivo": f"${h.get('total_extras_ef',0):,.0f}",
            "Recargo nocturno": f"${h.get('total_recargo',0):,.0f}",
            "Colaboradoras": len(h.get("colaboradores",[])),
        } for h in historico])
        st.dataframe(df_hist_res, hide_index=True, use_container_width=True)

        # Detalle de una quincena
        st.divider()
        st.subheader("Ver detalle de una quincena")
        opciones_q = [h.get("id","").replace("_"," al ") for h in historico]
        sel_q = st.selectbox("Selecciona quincena", opciones_q)
        if sel_q:
            q_data = next((h for h in historico
                           if h.get("id","").replace("_"," al ") == sel_q), None)
            if q_data:
                st.dataframe(pd.DataFrame([{
                    "Nombre": c.get("nombre",""),
                    "Horas": f"{c.get('tot_h',0):.1f}h",
                    "H. Extras": f"{c.get('ext_h',0):.1f}h",
                    "Recargo Noct.": f"${c.get('val_noct',0):,.0f}",
                    "Novedades": c.get("novedades_desc","—") or "—",
                    "Neto pagado": f"${c.get('neto',0):,.0f}",
                } for c in q_data.get("colaboradores",[])]),
                hide_index=True, use_container_width=True)

        # Histórico por colaboradora
        st.divider()
        st.subheader("Histórico por colaboradora")
        df_full = historico_a_dataframe()
        if not df_full.empty:
            colaboradoras = sorted(df_full["nombre"].unique())
            sel_col = st.selectbox("Colaboradora", colaboradoras)
            df_col = df_full[df_full["nombre"] == sel_col].copy()
            df_col["horas_fmt"] = df_col["horas_trabajadas"].apply(lambda x: f"{x:.1f}h")
            df_col["neto_fmt"]  = df_col["neto_pagado"].apply(lambda x: f"${x:,.0f}")
            st.dataframe(df_col[["periodo","horas_fmt","horas_extras",
                                   "val_recargo","neto_fmt","novedades"]].rename(columns={
                "periodo":"Período","horas_fmt":"Horas","horas_extras":"H. Extras",
                "val_recargo":"Recargo Noct.","neto_fmt":"Neto pagado","novedades":"Novedades"
            }), hide_index=True, use_container_width=True)

            # Exportar histórico
            buf_hist = io.BytesIO()
            df_col.to_excel(buf_hist, index=False)
            st.download_button(
                f"⬇️ Exportar histórico de {sel_col}",
                data=buf_hist.getvalue(),
                file_name=f"historico_{sel_col.replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# ═══════════════════════════════════════════════════════════════════════════════
# MÓDULO 4 — NÓMINA ELECTRÓNICA
# ═══════════════════════════════════════════════════════════════════════════════
elif "📊" in modulo:
    st.title("📊 Nómina electrónica mensual")
    st.caption("Genera el reporte mensual consolidando las dos quincenas — listo para Siigo.")

    historico = cargar_historico()
    colaboradores_db = cargar_colaboradores()

    if not historico:
        st.info("No hay quincenas guardadas. Procesa y guarda al menos una quincena primero.")
    else:
        # Seleccionar mes y año
        c1, c2 = st.columns(2)
        with c1:
            anios_disponibles = sorted(set(
                datetime.fromisoformat(h["periodo_ini"]).year
                for h in historico
            ), reverse=True)
            anio_sel = st.selectbox("Año", anios_disponibles)
        with c2:
            MESES_NOMBRES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
                             7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
            meses_con_data = sorted(set(
                datetime.fromisoformat(h["periodo_ini"]).month
                for h in historico
                if datetime.fromisoformat(h["periodo_ini"]).year == anio_sel
            ))
            if meses_con_data:
                mes_sel = st.selectbox("Mes", meses_con_data,
                                        format_func=lambda m: MESES_NOMBRES[m])
            else:
                st.warning("No hay quincenas guardadas para este año.")
                st.stop()

        # Verificar quincenas disponibles
        q_mes = [h for h in historico
                 if datetime.fromisoformat(h["periodo_ini"]).year == anio_sel
                 and datetime.fromisoformat(h["periodo_ini"]).month == mes_sel]
        q_mes.sort(key=lambda x: x["periodo_ini"])

        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            estado_q1 = "✅ Disponible" if len(q_mes) >= 1 else "❌ Falta"
            st.metric("Quincena 1 (1-15)", estado_q1,
                      delta=q_mes[0]["id"] if len(q_mes) >= 1 else None)
        with c2:
            estado_q2 = "✅ Disponible" if len(q_mes) >= 2 else "❌ Falta"
            st.metric("Quincena 2 (16-31)", estado_q2,
                      delta=q_mes[1]["id"] if len(q_mes) >= 2 else None)

        if len(q_mes) < 2:
            st.warning(f"Faltan quincenas para {MESES_NOMBRES[mes_sel]} {anio_sel}. "
                       "Procesa y guarda ambas quincenas del mes.")

        if st.button("📊 Generar nómina electrónica del mes", type="primary",
                     disabled=len(q_mes) < 1):
            with st.spinner("Generando Excel..."):
                # Preparar datos del mes
                datos_mes = preparar_datos_mes_desde_historico(
                    historico, mes_sel, anio_sel, colaboradores_db
                )
                meses_data = {mes_sel: datos_mes}
                xlsx_bytes = generar_nomina_electronica_xlsx(meses_data, anio_sel)

            nombre_mes = MESES_NOMBRES[mes_sel].upper()
            st.download_button(
                f"⬇️ Descargar NOMINA_ELECTRONICA_{nombre_mes}_{anio_sel}.xlsx",
                data=xlsx_bytes,
                file_name=f"NOMINA_ELECTRONICA_{nombre_mes}_{anio_sel}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.success("✅ Archivo generado. Ábrelo y pasa los datos a Siigo.")

        # Opción: generar todo el año
        st.divider()
        if st.button("📅 Generar nómina electrónica año completo"):
            with st.spinner("Generando Excel anual..."):
                meses_con_data_anio = sorted(set(
                    datetime.fromisoformat(h["periodo_ini"]).month
                    for h in historico
                    if datetime.fromisoformat(h["periodo_ini"]).year == anio_sel
                ))
                meses_data_anio = {}
                for m in meses_con_data_anio:
                    meses_data_anio[m] = preparar_datos_mes_desde_historico(
                        historico, m, anio_sel, colaboradores_db
                    )
                xlsx_anual = generar_nomina_electronica_xlsx(meses_data_anio, anio_sel)

            st.download_button(
                f"⬇️ NOMINA_ELECTRONICA_{anio_sel}_COMPLETA.xlsx",
                data=xlsx_anual,
                file_name=f"NOMINA_ELECTRONICA_{anio_sel}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
