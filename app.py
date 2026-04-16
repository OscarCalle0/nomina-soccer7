import streamlit as st
import pandas as pd
import os, io, zipfile
from datetime import datetime, date, timedelta

st.set_page_config(page_title="Nomina Soccer 7", page_icon="soccer", layout="wide")

from datos import (cargar_colaboradores, agregar_colaborador, actualizar_colaborador,
    retirar_colaborador, get_colaborador_por_reloj, cargar_historico,
    guardar_quincena_historico, historico_a_dataframe, inicializar_datos, Colaborador)
from colilla_pdf import generar_colilla_pdf, calcular_conceptos_colilla
from nomina_electronica import generar_nomina_electronica_xlsx, preparar_datos_mes_desde_historico
import sys
sys.path.insert(0, os.path.dirname(__file__))
from motor_nomina import procesar, calcular, crear_reporte_horarios, crear_resumen_nomina

inicializar_datos()

TIPOS_NOV = {"Sin novedad":None,"Vacaciones":"VACAC","Incapacidad EPS":"INC_EPS",
    "Incapacidad ARL":"INC_ARL","Licencia remunerada":"LIC_REM","Licencia no remunerada":"LIC_NREM",
    "Dia de la familia":"DIA_FAM","Dia compensatorio":"COMPENS","Calamidad":"CALAM",
    "Ausencia injustificada":"AUS_INJ","Suspension":"SUSPEND","Maternidad Paternidad":"MAT_PAT",
    "Renuncia Retiro":"RENUNCIA","Ingreso nuevo":"INGRESO"}
PCT_DEF = {"INC_EPS":66.66,"LIC_NREM":0.0,"SUSPEND":0.0,"AUS_INJ":0.0,"COMPENS":0.0}
DIAS_SEM = ["Lun","Mar","Mie","Jue","Vie","Sab","Dom"]

with st.sidebar:
    logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
    if os.path.exists(logo_path): st.image(logo_path, width=110)
    st.markdown("### Nomina Soccer 7")
    st.caption("GRANDA VARGAS SAS")
    st.divider()
    modulo = st.radio("Modulo:", ["Procesar quincena","Colaboradores","Historico","Nomina electronica"])

# ═══ MODULO 1 ═══════════════════════════════════════════════════
if modulo == "Procesar quincena":
    st.title("Procesar quincena")
    col_db_all = cargar_colaboradores()

    with st.expander("Paso 1 - Periodo", expanded=True):
        c1, c2 = st.columns(2)
        fecha_ini = c1.date_input("Desde", value=date.today().replace(day=1))
        fecha_fin = c2.date_input("Hasta", value=date.today().replace(day=15))
        st.info(f"{'Primera' if fecha_ini.day==1 else 'Segunda'} quincena de {fecha_ini.strftime('%B %Y')} - Meta 88h")

    with st.expander("Paso 2 - Archivo del reloj y revision", expanded=True):
        archivo = st.file_uploader("Archivo del Zkteco K50 (.xls .xlsx .csv)", type=["xls","xlsx","csv","txt"])
        if archivo:
            ext = archivo.name.lower().split(".")[-1]
            df = None
            try:
                if ext=="xls": df=pd.read_excel(archivo,engine="xlrd")
                elif ext in ["xlsx","xlsm"]: df=pd.read_excel(archivo,engine="openpyxl")
                else:
                    for sep in [",",";"]:
                        try: archivo.seek(0); tmp=pd.read_csv(archivo,sep=sep,encoding="utf-8-sig"); df=tmp if len(tmp.columns)>=4 else None; break
                        except: pass
                if df is None: st.error("No se pudo leer el archivo"); st.stop()
                df.columns=[str(c).strip() for c in df.columns]
                if "Numero" in df.columns: df.rename(columns={"Numero":"Numero"},inplace=True)
                p_ini=datetime.combine(fecha_ini,datetime.min.time())
                p_fin=datetime.combine(fecha_fin,datetime.max.time())
                aid=archivo.name+str(fecha_ini)+str(fecha_fin)
                if st.session_state.get("aid")!=aid:
                    res=procesar(df,p_ini,p_fin)
                    st.session_state.update({"df":df,"res":res,"p_ini":p_ini,"p_fin":p_fin,"novs_pi":{},"corrs":{},"aid":aid})
                    st.success(f"{archivo.name} - {len(df)} marcaciones - {len(res)} colaboradoras")
                else:
                    st.success(f"{archivo.name} cargado")
            except Exception as e: st.error(f"Error: {e}"); st.stop()

        res = st.session_state.get("res",[])
        if not res:
            st.info("Sube el archivo del reloj.")
        else:
            p_ini=st.session_state["p_ini"]; p_fin=st.session_state["p_fin"]
            novs_pi=st.session_state.get("novs_pi",{})
            th=sum(sum(d["trab"]*24 for d in e["dias"]) for e in res)
            ta=sum(len([d for d in e["dias"] if d["tiene"] and (d["ef"]=="missing" or d["sf"]=="missing")]) for e in res)
            c1,c2,c3,c4=st.columns(4)
            c1.metric("Colaboradoras",len(res)); c2.metric("Horas",f"{th:.0f}h")
            c3.metric("Alertas",ta); c4.metric("Sin reg",sum(len([d for d in e["dias"] if not d["tiene"]]) for e in res))
            st.divider()

            for ei,emp in enumerate(res):
                ht=sum(d["trab"]*24 for d in emp["dias"])
                na=len([d for d in emp["dias"] if d["tiene"] and (d["ef"]=="missing" or d["sf"]=="missing")])
                ns=len([d for d in emp["dias"] if not d["tiene"]])
                nt=sum(len(d.get("turnos",[])) for d in emp["dias"])
                db=get_colaborador_por_reloj(emp["nombre"],col_db_all)
                tit=f"{'ALERTA' if na else 'OK'} | {emp['nombre']} | {ht:.1f}h ({ht-88:+.1f}h) | {nt} turnos | {na} alertas | {ns} sin reg"
                if not db: tit+=" | NO EN BD"
                with st.expander(tit, expanded=bool(na)):
                    if not db: st.warning(f"{emp['nombre']} no esta en la base de datos")
                    for d in emp["dias"]:
                        dn=DIAS_SEM[d["fecha"].weekday()]; dk=d["dk"]
                        fl=f"{dn} {d['fecha'].day:02d}/{d['fecha'].month:02d}"
                        hd=d["trab"]*24; nd=d["noct"]*24
                        turnos=d.get("turnos",[])
                        if not d["tiene"]: st.write(f"**{fl}** Sin registro")
                        elif d["ef"]=="missing" or d["sf"]=="missing": st.write(f"**{fl}** ALERTA {hd:.2f}h {len(turnos)} turnos")
                        else: st.write(f"**{fl}** OK {hd:.2f}h {len(turnos)} turnos" + (f" Noct:{nd:.1f}h" if nd>0 else ""))

                        for ti,t in enumerate(turnos):
                            tt=(t["salida"]-t["entrada"]).total_seconds()/3600 if t["entrada"] and t["salida"] else 0
                            es=t["entrada"].strftime("%H:%M") if t["entrada"] else "FALTA"
                            ss=t["salida"].strftime("%H:%M") if t["salida"] else "FALTA"
                            x1,x2,x3,x4,x5,x6=st.columns([0.5,1,1,1.3,1.3,0.4])
                            x1.caption(f"T{ti+1}"); x2.caption(f"E:{es}"); x3.caption(f"S:{ss}")
                            ne=x4.text_input(f"ne {ei}{dk}{ti}",placeholder="HH:MM",key=f"ne_{ei}_{dk}_{ti}",label_visibility="collapsed")
                            ns2=x5.text_input(f"ns {ei}{dk}{ti}",placeholder="HH:MM",key=f"ns_{ei}_{dk}_{ti}",label_visibility="collapsed")
                            if tt>0: x6.caption(f"{tt:.1f}h")
                            if x6.button("X",key=f"d_{ei}_{dk}_{ti}"):
                                dfa=st.session_state["df"]; td=set()
                                if t["entrada"]: td.add(t["entrada"].strftime("%d/%m/%Y %H:%M:%S"))
                                if t["salida"]: td.add(t["salida"].strftime("%d/%m/%Y %H:%M:%S"))
                                if td:
                                    mask=~((dfa["Nombre"]==emp["nombre"])&(dfa["Tiempo"].astype(str).isin(td)))
                                    df2=dfa[mask].reset_index(drop=True)
                                    st.session_state["df"]=df2; st.session_state["res"]=procesar(df2,p_ini,p_fin); st.rerun()
                            if ne.strip() or ns2.strip():
                                st.session_state.setdefault("corrs",{})[f"{ei}_{dk}_{ti}"]={"nombre":emp["nombre"],"eid":emp["id"],"fdt":d["fecha"],"e":ne.strip() or None,"s":ns2.strip() or None}

                        a1,a2=st.columns(2)
                        ae=a1.text_input(f"ae{ei}{dk}",placeholder="Nueva entrada HH:MM",key=f"ae_{ei}_{dk}",label_visibility="collapsed")
                        aes=a2.text_input(f"as{ei}{dk}",placeholder="Nueva salida HH:MM",key=f"as_{ei}_{dk}",label_visibility="collapsed")
                        if ae.strip() or aes.strip():
                            st.session_state.setdefault("corrs",{})[f"add_{ei}_{dk}"]={"nombre":emp["nombre"],"eid":emp["id"],"fdt":d["fecha"],"e":ae.strip() or None,"s":aes.strip() or None}

                        nk=f"{emp['nombre']}||{dk}"; nav=novs_pi.get(nk,{}); nad=nav.get("desc","Sin novedad")
                        tl=list(TIPOS_NOV.keys()); idx=tl.index(nad) if nad in tl else 0
                        n1,n2=st.columns([3,1])
                        nsel=n1.selectbox(f"nov{ei}{dk}",tl,index=idx,key=f"nov_{ei}_{dk}",label_visibility="collapsed")
                        ndias=n2.number_input(f"nd{ei}{dk}",min_value=0.5,max_value=15.0,value=float(nav.get("dias",1.0)),step=0.5,key=f"nd_{ei}_{dk}",label_visibility="collapsed") if nsel!="Sin novedad" else 1.0
                        if nsel!="Sin novedad": novs_pi[nk]={"desc":nsel,"dias":ndias,"tipo":TIPOS_NOV[nsel]}
                        elif nk in novs_pi: del novs_pi[nk]
                        st.divider()

                    st.session_state["novs_pi"]=novs_pi
                    mc={k:v for k,v in st.session_state.get("corrs",{}).items() if v.get("nombre")==emp["nombre"]}
                    if mc:
                        st.info(f"{len(mc)} correccion(es) pendiente(s)")
                        if st.button(f"Aplicar correcciones {emp['nombre'].split()[0]}",key=f"ap_{ei}",type="primary"):
                            dfa=st.session_state["df"]; nf=[]; er=[]
                            for cv in mc.values():
                                fd=cv["fdt"]
                                if cv["e"]:
                                    try: h,m=map(int,cv["e"].split(":")); nf.append({"Numero":cv["eid"],"Nombre":cv["nombre"],"Tiempo":fd.replace(hour=h,minute=m,second=0).strftime("%d/%m/%Y %H:%M:%S"),"Estado":"Entrada","Dispositivos":"MANUAL","Tipo de Registro":0})
                                    except: er.append(f"Entrada invalida:{cv['e']}")
                                if cv["s"]:
                                    try:
                                        h,m=map(int,cv["s"].split(":")); ts=fd.replace(hour=h,minute=m,second=0)
                                        if h<6 and cv["e"] and int(cv["e"].split(":")[0])>=12: ts=(fd+timedelta(1)).replace(hour=h,minute=m,second=0)
                                        nf.append({"Numero":cv["eid"],"Nombre":cv["nombre"],"Tiempo":ts.strftime("%d/%m/%Y %H:%M:%S"),"Estado":"Salida","Dispositivos":"MANUAL","Tipo de Registro":0})
                                    except: er.append(f"Salida invalida:{cv['s']}")
                            if er:
                                for e in er: st.error(e)
                            else:
                                if nf: dfa=pd.concat([dfa,pd.DataFrame(nf)],ignore_index=True)
                                st.session_state["df"]=dfa; st.session_state["res"]=procesar(dfa,p_ini,p_fin)
                                for k in mc: del st.session_state["corrs"][k]
                                st.success("Correcciones aplicadas"); st.rerun()

                    m1,m2,m3,m4=st.columns(4)
                    m1.metric("Horas",f"{ht:.1f}h",delta=f"{ht-88:+.1f}h",delta_color="normal" if ht>=88 else "inverse")
                    m2.metric("Turnos",nt); m3.metric("Alertas",na); m4.metric("Sin reg",ns)

            # Consolidar novedades
            nd={e["nombre"]:[] for e in res}
            for nk,nv in novs_pi.items():
                nombre=nk.split("||")[0]; tc=nv.get("tipo")
                if tc and nombre in nd:
                    pct=PCT_DEF.get(tc,100.0)
                    if not any(n["tipo"]==tc for n in nd[nombre]):
                        nd[nombre].append({"tipo":tc,"dias":nv.get("dias",1.0),"pct":pct,"valor_override":None})
            st.session_state["novedades"]=nd

    with st.expander("Paso 3 - Calcular y descargar", expanded=True):
        if "res" not in st.session_state:
            st.info("Primero carga el archivo del reloj.")
        else:
            nd=st.session_state.get("novedades",{})
            nav=[(n,v) for n,v in nd.items() if v]
            if nav:
                for n,vs in nav:
                    for v in vs: st.caption(f"- {n}: {v['tipo']} {v['dias']}d {v['pct']:.0f}%")
                st.divider()

            if st.button("Calcular nomina", type="primary", use_container_width=True):
                p_ini=st.session_state["p_ini"]; p_fin=st.session_state["p_fin"]
                res=st.session_state["res"]; nd=st.session_state.get("novedades",{})
                rt=[]
                for emp in res:
                    db=get_colaborador_por_reloj(emp["nombre"],col_db_all)
                    sal=(db.salario_mensual if db.tipo=="empleado" else db.valor_hora_prestador) if db else 1423500
                    tp=db.tipo if db else "empleado"
                    t=calcular(emp,sal,tp,nd.get(emp["nombre"],[]))
                    rt.append((emp,t))
                st.session_state["rt"]=rt; st.success("Nomina calculada")

            if "rt" in st.session_state:
                rt=st.session_state["rt"]
                p_ini=st.session_state["p_ini"]; p_fin=st.session_state["p_fin"]
                rows=[]; gn=0
                for emp,t in rt:
                    db=get_colaborador_por_reloj(emp["nombre"],col_db_all)
                    if db and db.tipo=="prestador": neto=t["val_total_prest"]
                    else:
                        sal=db.salario_mensual if db else 1423500
                        sq=sal/2; di=t["dias_trab"]; aux=(249095/2)*di/15
                        ibc=sq+t["val_en"]+t["val_noct"]; ded=ibc*0.08+t["nov_deduccion"]
                        neto=sq+aux+t["val_en"]+t["val_noct"]+t["nov_devengado"]-ded
                    gn+=neto
                    rows.append({"Colaboradora":emp["nombre"],"Horas":f"{t['tot_h']:.1f}h","Extras":f"{t.get('en_h',0):.1f}h","Recargo":f"${t['val_noct']:,.0f}","Neto":f"${neto:,.0f}"})
                st.dataframe(pd.DataFrame(rows),hide_index=True,use_container_width=True)
                m1,m2,m3=st.columns(3)
                m1.metric("Total neto",f"${gn:,.0f}"); m2.metric("Extras efect",f"${sum(t.get('val_ee',0) for _,t in rt):,.0f}"); m3.metric("Recargo noct",f"${sum(t.get('val_noct',0) for _,t in rt):,.0f}")
                st.divider()
                tag=f"{p_ini.strftime('%Y%m%d')}_{p_fin.strftime('%Y%m%d')}"
                dc1,dc2,dc3=st.columns(3)
                with dc1:
                    if st.button("Generar colillas PDF",use_container_width=True):
                        lc=[]
                        for emp,t in rt:
                            db=get_colaborador_por_reloj(emp["nombre"],col_db_all)
                            cd={"id":db.id,"nombre_completo":db.nombre_completo,"cargo":db.cargo,"salario_mensual":db.salario_mensual,"tipo":db.tipo,"banco":db.banco,"cuenta":db.cuenta,"eps":db.eps,"valor_hora_prestador":db.valor_hora_prestador} if db else {"id":emp["id"],"nombre_completo":emp["nombre"],"cargo":"","salario_mensual":t["salario"],"tipo":t["tipo"],"banco":"","cuenta":"","eps":"","valor_hora_prestador":0}
                            lc.append(calcular_conceptos_colilla(t,cd))
                        st.session_state["pdf"]=generar_colilla_pdf(lc,p_ini,p_fin)
                    if st.session_state.get("pdf"):
                        st.download_button("Descargar colillas PDF",data=st.session_state["pdf"],file_name=f"COLILLAS_{tag}.pdf",mime="application/pdf",use_container_width=True)
                with dc2:
                    crear_reporte_horarios(rt,p_ini,p_fin,f"/tmp/rh_{tag}.xlsx")
                    st.download_button("Descargar reporte horarios",data=open(f"/tmp/rh_{tag}.xlsx","rb").read(),file_name=f"REPORTE_HORARIOS_{tag}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                with dc3:
                    vn=(1750905/2)+(249095/2)-(1750905/2)*0.08
                    crear_resumen_nomina(rt,vn,p_ini,p_fin,f"/tmp/rn_{tag}.xlsx")
                    st.download_button("Descargar resumen nomina",data=open(f"/tmp/rn_{tag}.xlsx","rb").read(),file_name=f"RESUMEN_NOMINA_{tag}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                st.divider()
                if st.button("Guardar en historico",type="primary"):
                    pid=f"{p_ini.strftime('%Y-%m-%d')}_{p_fin.strftime('%Y-%m-%d')}"
                    guardar_quincena_historico({"id":pid,"periodo_ini":p_ini.strftime("%Y-%m-%d"),"periodo_fin":p_fin.strftime("%Y-%m-%d"),"fecha_procesado":datetime.now().isoformat(),"total_nomina":gn,"total_extras_ef":sum(t.get("val_ee",0) for _,t in rt),"total_recargo":sum(t.get("val_noct",0) for _,t in rt),"colaboradores":[{"nombre":emp["nombre"],"tot_h":t.get("tot_h",0),"noct_h":t.get("noct_h",0),"ext_h":t.get("en_h",0)+t.get("ee_h",0),"en_h":t.get("en_h",0),"ee_h":t.get("ee_h",0),"val_en":t.get("val_en",0),"val_ee":t.get("val_ee",0),"val_noct":t.get("val_noct",0),"deu_h":t.get("deu_h",0),"dias_trab":t.get("dias_trab",0),"salario":t.get("salario",0),"tipo":t.get("tipo","empleado"),"val_total_prest":t.get("val_total_prest",0),"nov_devengado":t.get("nov_devengado",0),"nov_deduccion":t.get("nov_deduccion",0),"novedades_desc":" / ".join([nd2.get("desc","") for nd2 in t.get("nov_detalle",[])])} for emp,t in rt]})
                    st.success(f"Quincena {pid} guardada")

# ═══ MODULO 2 ═══════════════════════════════════════════════════
elif modulo == "Colaboradores":
    st.title("Colaboradores")
    cdb=cargar_colaboradores()
    t1,t2,t3=st.tabs(["Lista","Agregar","Editar Retirar"])
    with t1:
        st.dataframe(pd.DataFrame([{"Cedula":c.id,"Nombre":c.nombre_completo,"Reloj":c.nombre_reloj,"Cargo":c.cargo,"Salario":f"${c.salario_mensual:,.0f}","Tipo":c.tipo,"Ingreso":c.fecha_ingreso,"Retiro":c.fecha_retiro or "Activo","EPS":c.eps,"Banco":c.banco,"Cuenta":c.cuenta} for c in cdb]),hide_index=True,use_container_width=True)
    with t2:
        with st.form("fn"):
            c1,c2=st.columns(2)
            ced=c1.text_input("Cedula *"); nom=c1.text_input("Nombre completo *"); nr=c1.text_input("Nombre en reloj * (exacto)"); carg=c1.text_input("Cargo"); ing=c1.date_input("Fecha ingreso",value=date.today())
            tip=c2.selectbox("Tipo",["empleado","prestador"]); sal=c2.number_input("Salario",min_value=0,value=1750905,step=10000); vh=c2.number_input("Valor hora prestador",min_value=0,value=10000); ban=c2.text_input("Banco",value="AHORROS DAVIVIENDA"); cta=c2.text_input("Cuenta"); eps=c2.text_input("EPS",value="SAVIA SALUD")
            if st.form_submit_button("Agregar",type="primary"):
                if not ced or not nom or not nr: st.error("Cedula nombre y nombre en reloj son obligatorios")
                elif agregar_colaborador(Colaborador(id=ced.strip(),nombre_completo=nom.strip(),nombre_reloj=nr.strip(),cargo=carg.strip(),salario_mensual=sal,fecha_ingreso=ing.isoformat(),fecha_retiro=None,banco=ban.strip(),cuenta=cta.strip(),eps=eps.strip(),tipo=tip,valor_hora_prestador=vh if tip=="prestador" else 0,activo=True,notas="")): st.success(f"{nom} agregada"); st.rerun()
                else: st.error(f"Ya existe cedula {ced}")
    with t3:
        nms=[f"{c.nombre_completo} ({c.id})" for c in cdb]; sel=st.selectbox("Selecciona",nms)
        if sel:
            cs2=cdb[nms.index(sel)]
            with st.form("fe"):
                c1,c2=st.columns(2)
                nn=c1.text_input("Nombre completo",value=cs2.nombre_completo); nr2=c1.text_input("Nombre reloj",value=cs2.nombre_reloj); nc=c1.text_input("Cargo",value=cs2.cargo); ni=c1.date_input("Ingreso",value=date.fromisoformat(cs2.fecha_ingreso))
                ns2=c2.number_input("Salario",value=int(cs2.salario_mensual),step=10000); nb=c2.text_input("Banco",value=cs2.banco); ncu=c2.text_input("Cuenta",value=cs2.cuenta); ne2=c2.text_input("EPS",value=cs2.eps); nvh=c2.number_input("Valor hora",value=int(cs2.valor_hora_prestador))
                st.divider(); fr=st.date_input("Fecha retiro si aplica",value=date.fromisoformat(cs2.fecha_retiro) if cs2.fecha_retiro else date.today())
                b1,b2=st.columns(2)
                if b1.form_submit_button("Guardar",type="primary"):
                    cs2.nombre_completo=nn; cs2.nombre_reloj=nr2; cs2.cargo=nc; cs2.fecha_ingreso=ni.isoformat(); cs2.salario_mensual=ns2; cs2.banco=nb; cs2.cuenta=ncu; cs2.eps=ne2; cs2.valor_hora_prestador=nvh
                    actualizar_colaborador(cs2); st.success("Actualizado"); st.rerun()
                if b2.form_submit_button("Registrar retiro"):
                    retirar_colaborador(cs2.id,fr.isoformat()); st.success(f"Retiro {fr}"); st.rerun()

# ═══ MODULO 3 ═══════════════════════════════════════════════════
elif modulo == "Historico":
    st.title("Historico")
    h=cargar_historico()
    if not h: st.info("No hay quincenas guardadas.")
    else:
        m1,m2,m3=st.columns(3)
        m1.metric("Quincenas",len(h)); m2.metric("Total historico",f"${sum(x.get('total_nomina',0) for x in h):,.0f}"); m3.metric("Ultima",h[0].get("id","").replace("_"," al "))
        st.dataframe(pd.DataFrame([{"Periodo":x.get("id","").replace("_"," al "),"Procesado":x.get("fecha_procesado","")[:10],"Total":f"${x.get('total_nomina',0):,.0f}","Extras ef":f"${x.get('total_extras_ef',0):,.0f}","Recargo":f"${x.get('total_recargo',0):,.0f}"} for x in h]),hide_index=True,use_container_width=True)
        st.divider(); sel=st.selectbox("Ver detalle",[ x.get("id","").replace("_"," al ") for x in h])
        if sel:
            q=next((x for x in h if x.get("id","").replace("_"," al ")==sel),None)
            if q: st.dataframe(pd.DataFrame([{"Nombre":c.get("nombre",""),"Horas":f"{c.get('tot_h',0):.1f}h","Extras":f"{c.get('ext_h',0):.1f}h","Recargo":f"${c.get('val_noct',0):,.0f}","Novedades":c.get("novedades_desc","") or "ninguna"} for c in q.get("colaboradores",[])]),hide_index=True,use_container_width=True)
        df_full=historico_a_dataframe()
        if not df_full.empty:
            cl=sorted(df_full["nombre"].unique()); sc=st.selectbox("Historico por colaboradora",cl)
            dfc=df_full[df_full["nombre"]==sc]
            st.dataframe(dfc[["periodo","horas_trabajadas","horas_extras","val_recargo","neto_pagado","novedades"]].rename(columns={"periodo":"Periodo","horas_trabajadas":"Horas","horas_extras":"Extras","val_recargo":"Recargo","neto_pagado":"Neto","novedades":"Novedades"}),hide_index=True,use_container_width=True)
            buf=io.BytesIO(); dfc.to_excel(buf,index=False)
            st.download_button(f"Exportar historico {sc}",data=buf.getvalue(),file_name=f"hist_{sc.replace(' ','_')}.xlsx")

# ═══ MODULO 4 ═══════════════════════════════════════════════════
elif modulo == "Nomina electronica":
    st.title("Nomina electronica mensual")
    h=cargar_historico(); cdb=cargar_colaboradores()
    if not h: st.info("No hay quincenas guardadas.")
    else:
        MN={1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
        anios=sorted(set(datetime.fromisoformat(x["periodo_ini"]).year for x in h),reverse=True)
        c1,c2=st.columns(2); anio=c1.selectbox("Ano",anios)
        md=sorted(set(datetime.fromisoformat(x["periodo_ini"]).month for x in h if datetime.fromisoformat(x["periodo_ini"]).year==anio))
        if not md: st.warning("No hay datos."); st.stop()
        mes=c2.selectbox("Mes",md,format_func=lambda m:MN[m])
        qm=sorted([x for x in h if datetime.fromisoformat(x["periodo_ini"]).year==anio and datetime.fromisoformat(x["periodo_ini"]).month==mes],key=lambda x:x["periodo_ini"])
        c1,c2=st.columns(2); c1.metric("Quincena 1","OK" if len(qm)>=1 else "Falta"); c2.metric("Quincena 2","OK" if len(qm)>=2 else "Falta")
        if len(qm)<2: st.warning("Faltan quincenas del mes.")
        if st.button("Generar nomina electronica",type="primary",disabled=len(qm)<1):
            dm=preparar_datos_mes_desde_historico(h,mes,anio,cdb)
            xb=generar_nomina_electronica_xlsx({mes:dm},anio)
            st.download_button(f"Descargar NOMINA_{MN[mes].upper()}_{anio}.xlsx",data=xb,file_name=f"NOMINA_ELECTRONICA_{MN[mes].upper()}_{anio}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            st.success("Listo")
