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

# ── Estilos CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"] { background: #1B3A5C; }
[data-testid="stSidebar"] * { color: white !important; }
[data-testid="stSidebar"] .stRadio label { color: #cce0f5 !important; font-size: 14px; }
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { color: #93b4d4 !important; font-size: 12px; }
.metric-card { background: #f8f9fa; border-radius: 8px; padding: 12px 16px; border: 1px solid #e0e0e0; }
.badge-activo { background: #D5F5E3; color: #1E8449; padding: 2px 8px; border-radius: 12px; font-size: 12px; }
.badge-inactivo { background: #FADBD8; color: #922B21; padding: 2px 8px; border-radius: 12px; font-size: 12px; }
.badge-warn { background: #FFF3CD; color: #856404; padding: 2px 8px; border-radius: 12px; font-size: 12px; }
h1 { color: #1B3A5C !important; }
h2 { color: #1B3A5C !important; }
h3 { color: #2E6DA4 !important; font-size: 15px !important; }
</style>
""", unsafe_allow_html=True)

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
# ═══════════════════════════════════════════════════════════════════════════════
if "⚙️" in modulo:
    st.title("⚙️ Procesar quincena")
    colaboradores_db = cargar_colaboradores()

    # ── PASO 1: Período ─────────────────────────────────────────────────────
    with st.expander("📅 Paso 1 — Período de liquidación", expanded=True):
        c1, c2, c3 = st.columns([1, 1, 2])
        with c1:
            fecha_ini = st.date_input("Desde", value=date.today().replace(day=1))
        with c2:
            fecha_fin = st.date_input("Hasta", value=date.today().replace(day=15))
        with c3:
            q_num = "primera" if fecha_ini.day == 1 else "segunda"
            st.info(f"**{q_num.capitalize()} quincena** de {fecha_ini.strftime('%B %Y').capitalize()} · Meta: 88 horas")

    # ── PASO 2: Archivo del reloj ─────────────────────────────────────────
    with st.expander("📂 Paso 2 — Archivo del reloj biométrico", expanded=True):
        archivo = st.file_uploader(
            "Sube el reporte del Zkteco K50",
            type=["xls", "xlsx", "csv", "txt"],
            help="Acepta .xls, .xlsx y .csv exportados del reloj biométrico"
        )

        df_reloj = None
        if archivo:
            ext = archivo.name.lower().split(".")[-1]
            try:
                if ext == "xls":
                    df_reloj = pd.read_excel(archivo, engine="xlrd")
                elif ext in ["xlsx", "xlsm"]:
                    df_reloj = pd.read_excel(archivo, engine="openpyxl")
                else:
                    for sep in [",", ";", "\t"]:
                        try:
                            archivo.seek(0)
                            tmp = pd.read_csv(archivo, sep=sep, encoding="utf-8-sig")
                            if len(tmp.columns) >= 4:
                                df_reloj = tmp
                                break
                        except: pass

                if df_reloj is not None:
                    df_reloj.columns = [str(c).strip() for c in df_reloj.columns]
                    if "Numero" in df_reloj.columns and "Número" not in df_reloj.columns:
                        df_reloj.rename(columns={"Numero": "Número"}, inplace=True)

                    p_ini_dt = datetime.combine(fecha_ini, datetime.min.time())
                    p_fin_dt = datetime.combine(fecha_fin, datetime.max.time())
                    resultados_raw = procesar(df_reloj, p_ini_dt, p_fin_dt)

                    st.session_state["df_reloj"]       = df_reloj
                    st.session_state["resultados_raw"]  = resultados_raw
                    st.session_state["p_ini_dt"]        = p_ini_dt
                    st.session_state["p_fin_dt"]        = p_fin_dt
                    st.session_state["novedades"]       = {e["nombre"]: [] for e in resultados_raw}
                    st.session_state["marcaciones_manuales"] = []

                    st.success(f"✅ **{archivo.name}** · {len(df_reloj)} marcaciones · {len(resultados_raw)} colaboradoras")

                else:
                    st.error("No se pudo leer el archivo. Verifica el formato.")
            except Exception as e:
                st.error(f"Error leyendo el archivo: {e}")

    # ── PRE-INFORME con edición directa ───────────────────────────────────
    resultados_raw = st.session_state.get("resultados_raw", [])
    if resultados_raw:
        st.divider()
        st.markdown("### 📋 Pre-informe — Revisión y corrección")
        st.caption("Revisa cada colaboradora. Edita directamente las marcaciones incorrectas y aplica los cambios antes de calcular.")

        p_ini_dt = st.session_state["p_ini_dt"]
        p_fin_dt = st.session_state["p_fin_dt"]
        DIAS_SEM  = ["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"]

        # ── Métricas globales ──────────────────────────────────────────────
        total_alertas = sum(
            len([d for d in e["dias"] if d["tiene"] and (d["ef"]=="missing" or d["sf"]=="missing")])
            for e in resultados_raw)
        total_sin = sum(len([d for d in e["dias"] if not d["tiene"]]) for e in resultados_raw)
        total_ok  = sum(len([d for d in e["dias"] if d["tiene"] and d["ef"]=="ok" and d["sf"]=="ok"]) for e in resultados_raw)

        m1,m2,m3,m4 = st.columns(4)
        m1.metric("Colaboradoras", len(resultados_raw))
        m2.metric("Días completos ✅", total_ok)
        m3.metric("Alertas ⚠️", total_alertas, delta="requieren revisión" if total_alertas else None, delta_color="inverse")
        m4.metric("Sin registro ⚪", total_sin, delta="descanso o novedad" if total_sin else None, delta_color="off")

        st.divider()

        # ── Preinforme por colaboradora ────────────────────────────────────
        for emp_idx, emp in enumerate(resultados_raw):
            col_db     = get_colaborador_por_reloj(emp["nombre"], colaboradores_db)
            dias_ok    = [d for d in emp["dias"] if d["tiene"] and d["ef"]=="ok" and d["sf"]=="ok"]
            dias_alert = [d for d in emp["dias"] if d["tiene"] and (d["ef"]=="missing" or d["sf"]=="missing")]
            dias_sin   = [d for d in emp["dias"] if not d["tiene"]]
            horas_tot  = sum(d["trab"]*24 for d in emp["dias"] if d["trab"])

            icono  = "⚠️" if dias_alert else "✅"
            bd_txt = "✅ En BD" if col_db else "❌ No registrada en el sistema"
            titulo = (f"{icono} **{emp['nombre']}** — "
                      f"{horas_tot:.1f}h · {len(dias_ok)} OK · "
                      f"{len(dias_alert)} alertas · {len(dias_sin)} sin registro · {bd_txt}")

            with st.expander(titulo, expanded=bool(dias_alert)):
                if not col_db:
                    st.warning(f"⚠️ {emp['nombre']} no está en la base de datos. Agrégala en el módulo **Colaboradores** antes de continuar.")

                # Tabla de días con semáforo
                st.markdown("**Días del período:**")

                # Cabecera
                hc1,hc2,hc3,hc4,hc5,hc6,hc7 = st.columns([1.2,1.2,1.2,1.5,1.5,2,0.8])
                hc1.caption("Fecha"); hc2.caption("Entrada reloj"); hc3.caption("Salida reloj")
                hc4.caption("Nueva entrada"); hc5.caption("Nueva salida")
                hc6.caption("Estado"); hc7.caption("Horas")

                correcciones_emp = {}

                for d in emp["dias"]:
                    dia_nom   = DIAS_SEM[d["fecha"].weekday()]
                    fecha_str = f"{dia_nom} {d['fecha'].day:02d}/{d['fecha'].month:02d}"

                    entrada_reloj = d["entrada"].strftime("%H:%M") if d["entrada"] else "—"
                    salida_reloj  = d["salida"].strftime("%H:%M")  if d["salida"]  else "—"

                    if not d["tiene"]:
                        estado_txt = "⚪ Sin registro"
                        estado_color = "color:#888"
                    elif d["ef"] == "missing":
                        estado_txt = "🟠 Sin ENTRADA"
                        estado_color = "color:#E59866"
                    elif d["sf"] == "missing":
                        estado_txt = "🟠 Sin SALIDA"
                        estado_color = "color:#E59866"
                    else:
                        estado_txt = "🟢 Completo"
                        estado_color = "color:#1E8449"

                    horas_str = f"{d['trab']*24:.1f}h" if d["trab"] else "—"

                    row_key = f"pre_{emp_idx}_{d['dk']}"
                    c1,c2,c3,c4,c5,c6,c7 = st.columns([1.2,1.2,1.2,1.5,1.5,2,0.8])

                    with c1: st.markdown(f"<span style='font-size:13px'>{fecha_str}</span>", unsafe_allow_html=True)
                    with c2: st.markdown(f"<span style='font-family:monospace;font-size:13px'>{entrada_reloj}</span>", unsafe_allow_html=True)
                    with c3: st.markdown(f"<span style='font-family:monospace;font-size:13px'>{salida_reloj}</span>", unsafe_allow_html=True)
                    with c4:
                        nueva_entrada = st.text_input(
                            "NE", value="", placeholder="HH:MM",
                            key=f"ne_{emp_idx}_{d['dk']}",
                            label_visibility="collapsed"
                        )
                    with c5:
                        nueva_salida = st.text_input(
                            "NS", value="", placeholder="HH:MM",
                            key=f"ns_{emp_idx}_{d['dk']}",
                            label_visibility="collapsed"
                        )
                    with c6:
                        st.markdown(f"<span style='{estado_color};font-size:12px'>{estado_txt}</span>", unsafe_allow_html=True)
                    with c7:
                        st.markdown(f"<span style='font-size:12px;font-family:monospace'>{horas_str}</span>", unsafe_allow_html=True)

                    if nueva_entrada.strip() or nueva_salida.strip():
                        correcciones_emp[d["dk"]] = {
                            "nombre": emp["nombre"],
                            "fecha": d["fecha"].strftime("%d/%m/%Y"),
                            "entrada": nueva_entrada.strip() or None,
                            "salida":  nueva_salida.strip()  or None,
                        }

                # Botón aplicar correcciones de este empleado
                if correcciones_emp:
                    st.markdown("")
                    if st.button(f"✅ Aplicar correcciones de {emp['nombre'].split()[0]}",
                                 key=f"apply_{emp_idx}", type="primary"):
                        df_actual = st.session_state["df_reloj"]
                        filas_nuevas = []
                        errores_c = []

                        for dk, corr in correcciones_emp.items():
                            emp_data = next((e for e in resultados_raw if e["nombre"] == corr["nombre"]), None)
                            emp_id   = emp_data["id"] if emp_data else "0"
                            try:
                                fecha_dt = datetime.strptime(corr["fecha"], "%d/%m/%Y")
                            except:
                                errores_c.append(f"Fecha inválida: {corr['fecha']}"); continue

                            if corr["entrada"]:
                                try:
                                    hh,mm = map(int, corr["entrada"].split(":"))
                                    ts = fecha_dt.replace(hour=hh, minute=mm, second=0)
                                    filas_nuevas.append({
                                        "Número": emp_id, "Nombre": corr["nombre"],
                                        "Tiempo": ts.strftime("%d/%m/%Y %H:%M:%S"),
                                        "Estado": "Entrada", "Dispositivos": "MANUAL", "Tipo de Registro": 0
                                    })
                                except: errores_c.append(f"Hora entrada inválida: {corr['entrada']}")

                            if corr["salida"]:
                                try:
                                    hh,mm = map(int, corr["salida"].split(":"))
                                    if hh < 6 and corr["entrada"] and int(corr["entrada"].split(":")[0]) >= 12:
                                        ts_s = (fecha_dt + timedelta(days=1)).replace(hour=hh, minute=mm, second=0)
                                    else:
                                        ts_s = fecha_dt.replace(hour=hh, minute=mm, second=0)
                                    filas_nuevas.append({
                                        "Número": emp_id, "Nombre": corr["nombre"],
                                        "Tiempo": ts_s.strftime("%d/%m/%Y %H:%M:%S"),
                                        "Estado": "Salida", "Dispositivos": "MANUAL", "Tipo de Registro": 0
                                    })
                                except: errores_c.append(f"Hora salida inválida: {corr['salida']}")

                        if errores_c:
                            for e in errores_c: st.error(e)
                        elif filas_nuevas:
                            df_nuevo = pd.concat([df_actual, pd.DataFrame(filas_nuevas)], ignore_index=True)
                            nuevos_res = procesar(df_nuevo, p_ini_dt, p_fin_dt)
                            st.session_state["df_reloj"]      = df_nuevo
                            st.session_state["resultados_raw"] = nuevos_res
                            st.success(f"✅ {len(filas_nuevas)} marcación(es) aplicada(s) para {emp['nombre'].split()[0]}")
                            st.rerun()

                # Resumen rápido al final del expander
                st.divider()
                rc1, rc2, rc3 = st.columns(3)
                rc1.metric("Total horas", f"{horas_tot:.1f}h",
                           delta=f"{horas_tot-88:.1f}h vs meta" if horas_tot else None,
                           delta_color="normal")
                rc2.metric("Días trabajados", len(dias_ok) + len(dias_alert))
                rc3.metric("Días sin registro", len(dias_sin))


    with st.expander("✏️ Paso 3 — Marcaciones manuales (opcional)"):
        st.caption("Agrega días que el reloj no registró. Deja vacío si no aplica.")

        nombres_reloj = [e["nombre"] for e in st.session_state.get("resultados_raw", [])]
        if not nombres_reloj:
            st.info("Primero carga el archivo del reloj en el Paso 2.")
        else:
            if "marcaciones_form" not in st.session_state:
                st.session_state["marcaciones_form"] = [
                    {"nombre": "", "fecha": fecha_ini, "entrada": "08:00", "salida": "22:00"}
                ]

            mf = st.session_state["marcaciones_form"]

            for i, row in enumerate(mf):
                c1, c2, c3, c4, c5 = st.columns([2.5, 1.8, 1.2, 1.2, 0.4])
                with c1:
                    mf[i]["nombre"] = st.selectbox(
                        "Colaboradora", ["— Seleccionar —"] + nombres_reloj,
                        key=f"mn_{i}", index=nombres_reloj.index(row["nombre"])+1 if row["nombre"] in nombres_reloj else 0,
                        label_visibility="collapsed"
                    )
                with c2:
                    mf[i]["fecha"] = st.date_input("Fecha", value=row["fecha"],
                                                    min_value=fecha_ini, max_value=fecha_fin,
                                                    key=f"mf_{i}", label_visibility="collapsed")
                with c3:
                    mf[i]["entrada"] = st.text_input("Entrada", value=row["entrada"],
                                                      key=f"me_{i}", placeholder="HH:MM",
                                                      label_visibility="collapsed")
                with c4:
                    mf[i]["salida"] = st.text_input("Salida", value=row["salida"],
                                                     key=f"ms_{i}", placeholder="HH:MM",
                                                     label_visibility="collapsed")
                with c5:
                    if st.button("✕", key=f"mdel_{i}", help="Eliminar"):
                        mf.pop(i)
                        st.rerun()

            if st.button("➕ Agregar día manual"):
                mf.append({"nombre": "", "fecha": fecha_fin, "entrada": "08:00", "salida": "22:00"})
                st.rerun()

            if st.button("✅ Aplicar marcaciones manuales", type="primary"):
                df_actual = st.session_state.get("df_reloj")
                if df_actual is not None:
                    filas_nuevas = []
                    errores = []
                    resultados_raw = st.session_state.get("resultados_raw", [])

                    for row in mf:
                        if not row["nombre"] or row["nombre"] == "— Seleccionar —":
                            continue
                        emp_data = next((e for e in resultados_raw if e["nombre"] == row["nombre"]), None)
                        emp_id = emp_data["id"] if emp_data else "0"
                        fecha_dt = datetime.combine(row["fecha"], datetime.min.time())

                        try:
                            if row["entrada"]:
                                hh, mm = map(int, row["entrada"].split(":"))
                                ts = fecha_dt.replace(hour=hh, minute=mm, second=0)
                                filas_nuevas.append({"Número": emp_id, "Nombre": row["nombre"],
                                    "Tiempo": ts.strftime("%d/%m/%Y %H:%M:%S"),
                                    "Estado": "Entrada", "Dispositivos": "MANUAL", "Tipo de Registro": 0})
                        except: errores.append(f"Entrada inválida para {row['nombre']}")

                        try:
                            if row["salida"]:
                                hh, mm = map(int, row["salida"].split(":"))
                                if hh < 6 and row["entrada"] and int(row["entrada"].split(":")[0]) >= 12:
                                    ts_s = (fecha_dt + timedelta(days=1)).replace(hour=hh, minute=mm, second=0)
                                else:
                                    ts_s = fecha_dt.replace(hour=hh, minute=mm, second=0)
                                filas_nuevas.append({"Número": emp_id, "Nombre": row["nombre"],
                                    "Tiempo": ts_s.strftime("%d/%m/%Y %H:%M:%S"),
                                    "Estado": "Salida", "Dispositivos": "MANUAL", "Tipo de Registro": 0})
                        except: errores.append(f"Salida inválida para {row['nombre']}")

                    if errores:
                        for e in errores: st.error(e)
                    elif filas_nuevas:
                        df_combinado = pd.concat([df_actual, pd.DataFrame(filas_nuevas)], ignore_index=True)
                        p_ini_dt = st.session_state["p_ini_dt"]
                        p_fin_dt = st.session_state["p_fin_dt"]
                        nuevos_res = procesar(df_combinado, p_ini_dt, p_fin_dt)
                        st.session_state["df_reloj"] = df_combinado
                        st.session_state["resultados_raw"] = nuevos_res
                        st.success(f"✅ {len(filas_nuevas)} marcación(es) aplicada(s)")
                        st.rerun()

    # ── PASO 4: Novedades ─────────────────────────────────────────────────
    with st.expander("📋 Paso 4 — Novedades (incapacidades, vacaciones, etc.)"):
        st.caption("Registra novedades del período. Deja vacío si no hay.")

        TIPOS_NOV = {
            "INC_EPS":  "Incapacidad EPS",
            "INC_ARL":  "Incapacidad ARL (100%)",
            "MAT_PAT":  "Licencia maternidad/paternidad",
            "LIC_REM":  "Licencia remunerada",
            "LIC_NREM": "Licencia no remunerada",
            "VACAC":    "Vacaciones",
            "DIA_FAM":  "Día de la familia",
            "COMPENS":  "Día compensatorio",
            "CALAM":    "Calamidad doméstica",
            "SUSPEND":  "Suspensión disciplinaria",
            "AUS_INJ":  "Ausencia injustificada",
            "RENUNCIA": "Renuncia / Retiro (días trabajados)",
            "INGRESO":  "Ingreso nuevo (días trabajados en período)",
        }
        TIPOS_LISTA = list(TIPOS_NOV.values())
        TIPOS_KEYS  = list(TIPOS_NOV.keys())

        nombres_reloj = [e["nombre"] for e in st.session_state.get("resultados_raw", [])]
        if not nombres_reloj:
            st.info("Primero carga el archivo del reloj.")
        else:
            if "novedades_form" not in st.session_state:
                st.session_state["novedades_form"] = []

            nf = st.session_state["novedades_form"]

            if nf:
                c1h, c2h, c3h, c4h, c5h, _ = st.columns([2.5, 2.5, 0.8, 0.8, 1.5, 0.4])
                for col, txt in zip([c1h,c2h,c3h,c4h,c5h],
                                    ["Colaboradora","Tipo de novedad","Días","% Pago","Valor override"]):
                    col.caption(txt)

            for i, row in enumerate(nf):
                c1, c2, c3, c4, c5, c6 = st.columns([2.5, 2.5, 0.8, 0.8, 1.5, 0.4])
                with c1:
                    idx_n = nombres_reloj.index(row["nombre"]) if row["nombre"] in nombres_reloj else 0
                    nf[i]["nombre"] = st.selectbox("N", nombres_reloj, index=idx_n,
                                                    key=f"nn_{i}", label_visibility="collapsed")
                with c2:
                    idx_t = TIPOS_LISTA.index(row["tipo_desc"]) if row["tipo_desc"] in TIPOS_LISTA else 0
                    nf[i]["tipo_desc"] = st.selectbox("T", TIPOS_LISTA, index=idx_t,
                                                       key=f"nt_{i}", label_visibility="collapsed")
                    nf[i]["tipo"] = TIPOS_KEYS[TIPOS_LISTA.index(nf[i]["tipo_desc"])]
                with c3:
                    nf[i]["dias"] = st.number_input("D", min_value=0.5, max_value=30.0,
                                                     value=float(row.get("dias", 1)), step=0.5,
                                                     key=f"nd_{i}", label_visibility="collapsed")
                with c4:
                    pct_def = 100.0
                    if nf[i]["tipo"] in ["INC_EPS"]: pct_def = 66.66
                    if nf[i]["tipo"] in ["LIC_NREM","SUSPEND","AUS_INJ","COMPENS"]: pct_def = 0.0
                    nf[i]["pct"] = st.number_input("P", min_value=0.0, max_value=100.0,
                                                    value=float(row.get("pct", pct_def)), step=0.01,
                                                    key=f"np_{i}", label_visibility="collapsed")
                with c5:
                    vo = row.get("valor_override", None)
                    vo_str = str(int(vo)) if vo else ""
                    vo_input = st.text_input("V", value=vo_str, key=f"nv_{i}",
                                             placeholder="Opcional", label_visibility="collapsed")
                    nf[i]["valor_override"] = float(vo_input) if vo_input.strip() else None
                with c6:
                    if st.button("✕", key=f"ndel_{i}"):
                        nf.pop(i)
                        st.rerun()

            if st.button("➕ Agregar novedad"):
                nf.append({"nombre": nombres_reloj[0] if nombres_reloj else "",
                            "tipo": "VACAC", "tipo_desc": "Vacaciones",
                            "dias": 1.0, "pct": 100.0, "valor_override": None})
                st.rerun()

            # Guardar novedades en session_state
            novedades_dict = {n: [] for n in nombres_reloj}
            for row in nf:
                if row["nombre"] in novedades_dict:
                    novedades_dict[row["nombre"]].append({
                        "tipo": row["tipo"], "dias": row["dias"],
                        "pct": row["pct"], "valor_override": row["valor_override"]
                    })
            st.session_state["novedades"] = novedades_dict

    # ── PASO 5: Calcular y descargar ──────────────────────────────────────
    with st.expander("🚀 Paso 5 — Calcular y descargar", expanded=True):
        if "resultados_raw" not in st.session_state:
            st.info("Completa los pasos anteriores primero.")
        else:
            if st.button("🔄 Calcular nómina", type="primary", use_container_width=True):
                p_ini_dt = st.session_state["p_ini_dt"]
                p_fin_dt = st.session_state["p_fin_dt"]
                resultados_raw = st.session_state["resultados_raw"]
                novedades_dict = st.session_state.get("novedades", {})

                resultados_t = []
                for emp in resultados_raw:
                    col_db = get_colaborador_por_reloj(emp["nombre"], colaboradores_db)
                    if col_db:
                        sal  = col_db.salario_mensual if col_db.tipo == "empleado" else col_db.valor_hora_prestador
                        tipo = col_db.tipo
                    else:
                        sal, tipo = 1423500, "empleado"
                    novs = novedades_dict.get(emp["nombre"], [])
                    t = calcular(emp, sal, tipo, novs)
                    resultados_t.append((emp, t))

                st.session_state["resultados_t"] = resultados_t
                st.success("✅ Nómina calculada")

            if "resultados_t" in st.session_state:
                resultados_t = st.session_state["resultados_t"]
                p_ini_dt = st.session_state["p_ini_dt"]
                p_fin_dt = st.session_state["p_fin_dt"]

                # ── Tabla resumen ──────────────────────────────────────────
                st.subheader("Resumen de la quincena")
                rows_res = []
                gran_neto = 0
                for emp, t in resultados_t:
                    col_db = get_colaborador_por_reloj(emp["nombre"], colaboradores_db)
                    if col_db and col_db.tipo == "prestador":
                        neto = t["val_total_prest"]
                    else:
                        sal = col_db.salario_mensual if col_db else 1423500
                        sal_q = sal / 2
                        dias = t["dias_trab"]
                        aux  = (249095/2) * dias / 15
                        ibc  = sal_q + t["val_en"] + t["val_noct"]
                        ded  = ibc * 0.08 + t["nov_deduccion"]
                        neto = sal_q + aux + t["val_en"] + t["val_noct"] + t["nov_devengado"] - ded

                    gran_neto += neto
                    estado = (f"✅ +{t['en_h']:.1f}h extras" if t.get("en_h",0) > 0
                              else f"⚠️ Debe {t['deu_h']:.1f}h" if t.get("deu_h",0) > 0.1
                              else "✓ OK")
                    novedades_txt = ", ".join([
                        nd.get("desc","").split(" —")[0] for nd in t.get("nov_detalle",[])
                    ]) or "—"
                    rows_res.append({
                        "Colaboradora": emp["nombre"],
                        "Horas": f"{t['tot_h']:.1f}h",
                        "H. Extras": f"{t.get('en_h',0):.1f}h",
                        "Recargo Noct.": f"${t['val_noct']:,.0f}",
                        "Novedades": novedades_txt,
                        "Estado": estado,
                        "Neto a pagar": f"${neto:,.0f}",
                    })

                st.dataframe(pd.DataFrame(rows_res), hide_index=True, use_container_width=True)

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total neto", f"${gran_neto:,.0f}")
                extras_ef = sum(t.get("val_ee",0) for _,t in resultados_t)
                m2.metric("Extras efectivo", f"${extras_ef:,.0f}")
                recargos  = sum(t.get("val_noct",0) for _,t in resultados_t)
                m3.metric("Recargo nocturno", f"${recargos:,.0f}")
                m4.metric("Colaboradoras", len(resultados_t))

                st.divider()
                st.subheader("Descargar archivos")

                # ── Generar archivos ───────────────────────────────────────
                tag = f"{p_ini_dt.strftime('%Y%m%d')}_{p_fin_dt.strftime('%Y%m%d')}"

                col1, col2, col3, col4 = st.columns(4)

                # 1. Colillas PDF
                with col1:
                    if st.button("📄 Generar colillas PDF", use_container_width=True):
                        lista_colillas = []
                        for emp, t in resultados_t:
                            col_db = get_colaborador_por_reloj(emp["nombre"], colaboradores_db)
                            if col_db:
                                col_dict = {
                                    "id": col_db.id, "nombre_completo": col_db.nombre_completo,
                                    "cargo": col_db.cargo, "salario_mensual": col_db.salario_mensual,
                                    "tipo": col_db.tipo, "banco": col_db.banco,
                                    "cuenta": col_db.cuenta, "eps": col_db.eps,
                                    "valor_hora_prestador": col_db.valor_hora_prestador,
                                }
                            else:
                                col_dict = {
                                    "id": emp["id"], "nombre_completo": emp["nombre"],
                                    "cargo": "", "salario_mensual": t["salario"],
                                    "tipo": t["tipo"], "banco": "", "cuenta": "", "eps": "",
                                    "valor_hora_prestador": 0,
                                }
                            datos_colilla = calcular_conceptos_colilla(t, col_dict)
                            lista_colillas.append(datos_colilla)

                        pdf_bytes = generar_colilla_pdf(lista_colillas, p_ini_dt, p_fin_dt)
                        st.download_button(
                            "⬇️ Descargar colillas PDF",
                            data=pdf_bytes,
                            file_name=f"COLILLAS_{tag}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )

                # 2. Reporte horarios Excel
                with col2:
                    buf_rh = io.BytesIO()
                    resultados_t_calc = []
                    for emp, t in resultados_t:
                        resultados_t_calc.append((emp, t))
                    crear_reporte_horarios(resultados_t_calc, p_ini_dt, p_fin_dt,
                                           f"/tmp/rh_{tag}.xlsx")
                    with open(f"/tmp/rh_{tag}.xlsx", "rb") as f:
                        rh_bytes = f.read()
                    st.download_button(
                        "⬇️ Reporte horarios",
                        data=rh_bytes,
                        file_name=f"REPORTE_HORARIOS_{tag}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                # 3. Resumen nómina Excel
                with col3:
                    vd_sal = 1750905
                    vd_q = vd_sal / 2
                    vd_aux = 249095 / 2
                    vd_ibc = vd_q
                    vd_ded = vd_ibc * 0.08
                    valentina_neto = vd_q + vd_aux - vd_ded

                    crear_resumen_nomina(resultados_t_calc, valentina_neto, p_ini_dt, p_fin_dt,
                                         f"/tmp/rn_{tag}.xlsx")
                    with open(f"/tmp/rn_{tag}.xlsx", "rb") as f:
                        rn_bytes = f.read()
                    st.download_button(
                        "⬇️ Resumen nómina",
                        data=rn_bytes,
                        file_name=f"RESUMEN_NOMINA_{tag}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                # 4. ZIP con todo
                with col4:
                    if st.button("📦 Descargar todo (ZIP)", use_container_width=True):
                        zip_buf = io.BytesIO()
                        with zipfile.ZipFile(zip_buf, "w") as zf:
                            zf.writestr(f"COLILLAS_{tag}.pdf", pdf_bytes if "pdf_bytes" in dir() else b"")
                            zf.write(f"/tmp/rh_{tag}.xlsx", f"REPORTE_HORARIOS_{tag}.xlsx")
                            zf.write(f"/tmp/rn_{tag}.xlsx", f"RESUMEN_NOMINA_{tag}.xlsx")
                        st.download_button(
                            "⬇️ Descargar ZIP",
                            data=zip_buf.getvalue(),
                            file_name=f"NOMINA_COMPLETA_{tag}.zip",
                            mime="application/zip",
                            use_container_width=True,
                        )

                # ── Guardar en histórico ───────────────────────────────────
                st.divider()
                if st.button("💾 Guardar esta quincena en el histórico", type="primary"):
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
                            {
                                "nombre": emp["nombre"],
                                "tot_h": t.get("tot_h",0), "noct_h": t.get("noct_h",0),
                                "ext_h": t.get("en_h",0)+t.get("ee_h",0),
                                "en_h": t.get("en_h",0), "ee_h": t.get("ee_h",0),
                                "val_en": t.get("val_en",0), "val_ee": t.get("val_ee",0),
                                "val_noct": t.get("val_noct",0), "val_deu": t.get("val_deu",0),
                                "deu_h": t.get("deu_h",0), "dias_trab": t.get("dias_trab",0),
                                "salario": t.get("salario",0), "tipo": t.get("tipo","empleado"),
                                "val_total_prest": t.get("val_total_prest",0),
                                "nov_devengado": t.get("nov_devengado",0),
                                "nov_deduccion": t.get("nov_deduccion",0),
                                "novedades_desc": " / ".join([
                                    nd.get("desc","") for nd in t.get("nov_detalle",[])
                                ]),
                                "neto": next(
                                    (neto for e2,t2 in resultados_t for neto in
                                     [t2.get("val_total_prest",0) if t2.get("tipo")=="prestador"
                                      else (t2.get("salario",0)/2*(t2.get("dias_trab",15)/15)
                                            + (249095/2)*(t2.get("dias_trab",15)/15)
                                            + t2.get("val_en",0)+t2.get("val_noct",0)
                                            +t2.get("nov_devengado",0)
                                            -(t2.get("salario",0)/2*(t2.get("dias_trab",15)/15)
                                              +t2.get("val_en",0)+t2.get("val_noct",0))*0.08
                                            -t2.get("nov_deduccion",0)
                                      )]
                                    if e2["nombre"] == emp["nombre"]),
                                    0
                                ),
                            }
                            for emp, t in resultados_t
                        ],
                    }
                    guardar_quincena_historico(registro)
                    st.success(f"✅ Quincena {periodo_id} guardada en el histórico")


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
