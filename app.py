"""
Kumara  -  Cotizador de eventos de emergencias
Misma lógica del Excel "Cotizador Kumara", pero modular y en la web:
  - Configuración (editable): parámetros, pistas y recursos, menús, vehículos.
  - Cotización (se diligencia por evento): datos, pistas, alimentación, transporte.
  - Resumen con exportación a Word y PDF (sin macros).
"""
import json
import os
from datetime import date, datetime

import pandas as pd
import streamlit as st

from exportar import generar_pdf, generar_word, nombre_archivo
from modelos import (CAFE, REGLAS, REGLAS_AYUDA, SERVICIOS_ALIMENTACION, NO, SI,
                     armar_cotizacion, calcular_alimentacion, calcular_linea,
                     calcular_transporte, cant, cantidad_requerida, config_defecto,
                     cop, evento_defecto, fusionar, num, ov_pista, ov_recurso,
                     pista_por_nombre, slug, sugerido, txt, valor_base)

RUTA_CONFIG = "config_kumara.json"      # si existe en el repositorio, es la configuración inicial
LOGO = os.path.join("images", "logo_kumara.png")

st.set_page_config(page_title="Kumara - Cotizador de eventos", page_icon="🧯",
                   layout="wide", initial_sidebar_state="expanded")

# =============================================================================
#  ESTILO  (paleta tomada del Excel: azul 193657 · turquesa 087E8B)
# =============================================================================
st.markdown("""
<style>
.stApp { background-color: #F4F7FA; }
section[data-testid="stSidebar"] { background-color: #193657 !important; }
section[data-testid="stSidebar"] * { color: #E8EEF4 !important; }
h1 { color: #193657 !important; }
h2, h3 { color: #087E8B !important; }
.stButton > button, .stDownloadButton > button {
    background-color: #087E8B !important; color: #FFFFFF !important;
    border: none !important; border-radius: 6px !important; font-weight: 600 !important;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    background-color: #193657 !important; color: #FFFFFF !important;
}
section[data-testid="stSidebar"] .stButton > button,
section[data-testid="stSidebar"] .stDownloadButton > button {
    background-color: #087E8B !important; color: #FFFFFF !important; width: 100%;
}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] section {
    background-color: #0F2640 !important; border: 1px dashed #5FB3BC !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {
    background-color: #087E8B !important; color: #FFFFFF !important;
}
[data-testid="stMetric"] {
    background: #DDF5E5; border: 1px solid #087E8B; border-radius: 8px; padding: 8px 12px;
}
.info-box { background: #DCEEFF; border-left: 4px solid #087E8B; padding: 10px 14px;
            border-radius: 4px; margin-bottom: 12px; font-size: .92em; color: #193657; }
.aviso { background: #FFF0B3; border-left: 4px solid #C9A400; padding: 8px 14px;
         border-radius: 4px; margin: 4px 0; font-size: .9em; color: #4A3F00; }
.total-box { background: #193657; color: #FFFFFF; border-radius: 10px; padding: 14px 20px;
             text-align: center; font-size: 1.4em; font-weight: 700; margin: 10px 0; }
.total-box small { display: block; font-size: .55em; font-weight: 400; color: #9FD3D9; }
.paso { background: #FFFFFF; border: 1.5px solid #087E8B; border-radius: 10px;
        padding: 10px 12px; text-align: center; font-size: .9em; color: #193657; }
.paso b { color: #087E8B; }
.paso.cfg { border-color: #193657; background: #DCEEFF; }
.cab { font-size: .78em; font-weight: 700; color: #193657; }
.pie { text-align: center; color: #6B7C8F; font-size: .8em; margin-top: 28px;
       padding-top: 8px; border-top: 1px solid #C9D3DC; }
</style>
""", unsafe_allow_html=True)


# =============================================================================
#  ESTADO
# =============================================================================
def cargar_config_repo():
    if os.path.exists(RUTA_CONFIG):
        try:
            with open(RUTA_CONFIG, encoding="utf-8") as f:
                return fusionar(config_defecto(), json.load(f))
        except Exception:
            return None
    return None


def init():
    ss = st.session_state
    if "config" not in ss:
        ss.config = cargar_config_repo() or config_defecto()
    if "evento" not in ss:
        ss.evento = evento_defecto(ss.config)
    ss.setdefault("ver", 0)          # versión de los widgets del evento
    ss.setdefault("cfg_ver", 0)      # versión de los editores de configuración
    ss.setdefault("snap", {})        # copias base de las tablas editables
    ss.setdefault("pedir_nueva", False)
    ss.setdefault("mensaje", "")


init()


def CFG():
    return st.session_state.config


def EV():
    return st.session_state.evento


def k(*partes):
    """Clave de widget única y ligada a la versión actual del evento."""
    return f"v{st.session_state.ver}_" + "_".join(slug(p) for p in partes)


def info(texto):
    st.markdown(f'<div class="info-box">{texto}</div>', unsafe_allow_html=True)


def aviso(texto):
    st.markdown(f'<div class="aviso">⚠ {texto}</div>', unsafe_allow_html=True)


def cabeceras(cols, titulos):
    for c, t in zip(cols, titulos):
        c.markdown(f'<span class="cab">{t}</span>', unsafe_allow_html=True)


def entero(v):
    return int(round(num(v)))


# =============================================================================
#  INICIO / GUÍA
# =============================================================================
def pagina_inicio():
    st.title("Kumara")
    st.markdown("**Cotizador de eventos de emergencias** — pistas, alimentación y transporte "
                "en una sola cotización, lista para enviar en Word o PDF.")

    st.subheader("Cómo se usa")
    st.markdown("**A. Configuración** — se hace una vez y se reutiliza en todas las cotizaciones.")
    cols = st.columns(4)
    for c, (t, d) in zip(cols, [
            ("Parámetros", "tarifa del profesional, cuántos asistentes por profesional, café, IVA"),
            ("Pistas y recursos", "qué lleva cada pista, cuántas personas atiende cada recurso y su valor"),
            ("Menús", "opciones y tarifas de cada servicio de alimentación"),
            ("Vehículos", "capacidad y tarifa de aerovan, taxi u otros")]):
        c.markdown(f'<div class="paso cfg"><b>{t}</b><br>{d}</div>', unsafe_allow_html=True)

    st.markdown("**B. Cotización** — se diligencia por cada evento, en este orden.")
    cols = st.columns(5)
    for c, (n, t, d) in zip(cols, [
            ("1", "Datos del evento", "empresa, asistentes, horario"),
            ("2", "Pistas", "elige las pistas y ajusta cada recurso"),
            ("3", "Alimentación", "servicios, menús y domicilio"),
            ("4", "Transporte", "aerovans y taxis según asistentes"),
            ("5", "Resumen", "totales, IVA opcional, Word y PDF")]):
        c.markdown(f'<div class="paso"><b>{n}. {t}</b><br>{d}</div>', unsafe_allow_html=True)

    st.subheader("Conceptos clave")
    st.markdown("""
- **Requerido**: lo que exige el evento según los asistentes y la regla del recurso.
- **Aporta empresa**: lo que el cliente ya tiene. **A cotizar = Requerido − Aporta empresa**.
- **Incluir**: desmarca un recurso para sacarlo de la cotización sin borrarlo.
- **Valor unitario**: viene del catálogo, pero lo puedes cambiar solo para este evento.
- **Recurso adicional**: en cada pista puedes agregar un ítem que aplica solo a este evento.
- **Alimentación**: «Sugerido» se calcula con el horario; tú decides si lo incluyes.
""")
    st.markdown('<div class="aviso">Los datos viven en tu navegador. '
                '<b>Guarda la cotización (.json) antes de cerrar la pestaña</b> y '
                'cárgala cuando quieras retomarla. Lo mismo aplica a la configuración '
                'si la modificas.</div>', unsafe_allow_html=True)


# =============================================================================
#  1. DATOS DEL EVENTO
# =============================================================================
def pagina_evento():
    cfg, ev = CFG(), EV()
    st.title("1. Datos del evento")
    info("Con estos datos se calculan las cantidades de todas las pistas.")

    c1, c2 = st.columns(2)
    ev["empresa"] = c1.text_input("Empresa", value=ev["empresa"], key=k("empresa"))
    ev["contacto"] = c2.text_input("Contacto", value=ev["contacto"], key=k("contacto"))
    c1, c2, c3 = st.columns(3)
    ev["ciudad"] = c1.text_input("Ciudad", value=ev["ciudad"], key=k("ciudad"))
    try:
        f0 = date.fromisoformat(ev["fecha"]) if ev["fecha"] else None
    except ValueError:
        f0 = None
    f = c2.date_input("Fecha del evento", value=f0, format="DD/MM/YYYY", key=k("fecha"))
    ev["fecha"] = f.isoformat() if f else ""
    niveles = cfg["parametros"]["niveles_brigada"] or ["-"]
    idx = niveles.index(ev["nivel"]) if ev["nivel"] in niveles else 0
    ev["nivel"] = c3.selectbox("Nivel de brigada", niveles, index=idx, key=k("nivel"))
    c1, c2 = st.columns(2)
    ev["lugar"] = c1.text_input("Lugar del evento", value=ev["lugar"], key=k("lugar"))
    ev["direccion"] = c2.text_input("Dirección", value=ev["direccion"], key=k("direccion"))

    st.subheader("Asistentes y horario")
    c1, c2, c3, c4 = st.columns(4)
    ev["asistentes"] = c1.number_input("Número de asistentes", min_value=0, step=1,
                                       value=entero(ev["asistentes"]), key=k("asistentes"))
    ev["h_inicio"] = c2.number_input("Hora de inicio (0-24)", min_value=0.0, max_value=24.0,
                                     step=0.5, format="%.1f", value=float(num(ev["h_inicio"])),
                                     key=k("h_inicio"))
    ev["h_fin"] = c3.number_input("Hora de fin (0-24)", min_value=0.0, max_value=24.0,
                                  step=0.5, format="%.1f", value=float(num(ev["h_fin"])),
                                  key=k("h_fin"))
    dur = max(0.0, num(ev["h_fin"]) - num(ev["h_inicio"]))
    c4.metric("Duración (horas)", f"{dur:g}")

    params = cfg["parametros"]
    prof = -(-entero(ev["asistentes"]) // max(1, entero(params["prof_por_asistentes"])))
    st.metric("Profesionales sugeridos por pista",
              f"{prof}", help=f"1 profesional por cada {entero(params['prof_por_asistentes'])} asistentes")
    if num(ev["h_fin"]) <= num(ev["h_inicio"]):
        aviso("La hora de fin debe ser mayor que la hora de inicio.")


# =============================================================================
#  2. PISTAS
# =============================================================================
def descripcion_regla(rec, params):
    r = rec.get("regla")
    if r == "fija":
        return "fija"
    if r == "profesionales":
        return f"1 por {entero(params['prof_por_asistentes'])}"
    if r == "proporcional":
        return f"{rec.get('ref', '?')} ÷ {cant(rec.get('divisor'))}"
    p = num(rec.get("personas"))
    return f"1 por {cant(p)}" if p > 0 else "-"


def render_pista(cfg, ev, nombre):
    pista = pista_por_nombre(cfg, nombre)
    datos = ov_pista(ev, nombre)
    asist = num(ev["asistentes"])
    params = cfg["parametros"]

    anchos = [3.0, 1.2, 1.0, 1.2, 1.0, 1.5, 1.5, 0.8, 2.2]
    cabeceras(st.columns(anchos), ["Recurso", "Regla", "Requerido", "Aporta empresa",
                                   "A cotizar", "Valor unitario", "Subtotal", "Incluir",
                                   "Observaciones"])
    cache, total = {}, 0.0
    for rec in pista["recursos"]:
        o = ov_recurso(ev, nombre, rec["nombre"])
        cols = st.columns(anchos, vertical_alignment="center")
        unidad = rec.get("unidad") or ""
        cols[0].write(rec["nombre"] + (f" ({unidad})" if unidad and unidad != "unidad" else ""))
        cols[1].caption(descripcion_regla(rec, params))

        o["incluir"] = cols[7].checkbox("Incluir", value=bool(o["incluir"]),
                                        key=k(nombre, rec["nombre"], "inc"),
                                        label_visibility="collapsed")
        o["aporta"] = cols[3].number_input("Aporta empresa", min_value=0, step=1,
                                           value=entero(o["aporta"]),
                                           key=k(nombre, rec["nombre"], "apo"),
                                           label_visibility="collapsed")
        vb = entero(valor_base(rec, params))
        actual = entero(o["valor"]) if o["valor"] is not None else vb
        val = cols[5].number_input("Valor unitario", min_value=0, step=1000, value=actual,
                                   key=k(nombre, rec["nombre"], "val"),
                                   label_visibility="collapsed")
        o["valor"] = None if val == vb else val
        o["obs"] = cols[8].text_input("Obs.", value=o["obs"],
                                      key=k(nombre, rec["nombre"], "obs"),
                                      label_visibility="collapsed")

        lin = calcular_linea(rec, o, asist, params, cache, True)
        cache[rec["nombre"]] = lin["a_cotizar"]
        total += lin["subtotal"]
        cols[2].write(cant(lin["requerida"]))
        cols[4].write(cant(lin["a_cotizar"]))
        cols[6].write(cop(lin["subtotal"]))

    extras = datos.setdefault("extras", [])
    for i, ex in enumerate(list(extras)):
        cols = st.columns(anchos, vertical_alignment="center")
        c, v = num(ex.get("cantidad")), num(ex.get("valor"))
        cols[0].write("➕ " + ex.get("nombre", ""))
        cols[1].caption("adicional")
        cols[2].write(cant(c))
        cols[4].write(cant(c))
        cols[5].write(cop(v))
        cols[6].write(cop(c * v))
        cols[8].caption(ex.get("obs", ""))
        if cols[7].button("✕", key=k(nombre, "ex_del", i), help="Quitar este recurso adicional"):
            extras.pop(i)
            st.rerun()
        total += c * v

    with st.expander("Agregar un recurso adicional (solo para este evento)"):
        n = len(extras)
        c1, c2, c3, c4 = st.columns([3, 1, 1.5, 2.5])
        nom = c1.text_input("Nombre", key=k(nombre, "ex_nom", n))
        can = c2.number_input("Cantidad", min_value=0, step=1, key=k(nombre, "ex_can", n))
        valor = c3.number_input("Valor unitario", min_value=0, step=1000,
                                key=k(nombre, "ex_val", n))
        obs = c4.text_input("Observación", key=k(nombre, "ex_obs", n))
        if st.button("Agregar recurso", key=k(nombre, "ex_add", n)):
            if nom.strip() and can > 0:
                extras.append({"nombre": nom.strip(), "cantidad": can,
                               "valor": valor, "obs": obs})
                st.rerun()
            else:
                st.warning("Escribe un nombre y una cantidad mayor que cero.")

    st.markdown(f'<div class="total-box"><small>Subtotal {nombre}</small>{cop(total)}</div>',
                unsafe_allow_html=True)


def pagina_pistas():
    cfg, ev = CFG(), EV()
    st.title("2. Pistas")
    info("Elige las pistas del evento. Para cada recurso: <b>Requerido</b> sale de los asistentes; "
         "si el cliente ya tiene parte, anótalo en <b>Aporta empresa</b>.")
    if num(ev["asistentes"]) <= 0:
        aviso("Primero ingresa el número de asistentes en «1. Datos del evento».")

    nombres = [p["nombre"] for p in cfg["pistas"]]
    activas = [n for n in nombres if ev["pistas"].get(n, {}).get("activa")]
    sel = st.multiselect("Pistas incluidas en este evento", nombres, default=activas,
                         key=k("pistas_sel"))
    for n in nombres:
        ov_pista(ev, n)["activa"] = n in sel
    if not sel:
        st.info("Selecciona al menos una pista para empezar.")
        return
    for tab, nombre in zip(st.tabs(sel), sel):
        with tab:
            render_pista(cfg, ev, nombre)


# =============================================================================
#  3. ALIMENTACIÓN
# =============================================================================
def pagina_alimentacion():
    cfg, ev = CFG(), EV()
    st.title("3. Alimentación")
    info("«Sugerido» se calcula con el horario del evento, pero nada se incluye hasta que "
         "marques <b>Incluir</b>. Deja el valor editable en 0 para usar la tarifa del menú.")
    alim = ev.setdefault("alimentacion", {})

    anchos = [1.7, 0.9, 0.8, 2.4, 1.2, 1.2, 1.4, 1.3, 1.5]
    cabeceras(st.columns(anchos), ["Servicio", "Sugerido", "Incluir", "Menú",
                                   "Personas (0 = asistentes)", "Tarifa",
                                   "Valor editable (0 = tarifa)", "Domicilio", "Total"])
    filas = []
    for serv in SERVICIOS_ALIMENTACION + [CAFE]:
        d = alim.setdefault(serv, {"incluir": False, "menu": "", "personas": 0,
                                   "valor": 0, "domicilio": 0})
        cols = st.columns(anchos, vertical_alignment="center")
        cols[0].write(f"**{serv}**")
        if serv == CAFE:
            cols[1].write("-")
        else:
            sug = sugerido(cfg, serv, ev["h_inicio"], ev["h_fin"])
            cols[1].write(f"{'✅' if sug == SI else '·'} {sug}")
        d["incluir"] = cols[2].checkbox("Incluir", value=bool(d["incluir"]),
                                        key=k("al", serv, "inc"), label_visibility="collapsed")
        if serv == CAFE:
            por = entero(cfg["cafe"]["personas_por_estacion"])
            cols[3].caption(f"1 estación cada {por} personas")
        else:
            opciones = [m["nombre"] for m in cfg["menus"].get(serv, [])]
            if opciones:
                i = opciones.index(d["menu"]) if d["menu"] in opciones else 0
                d["menu"] = cols[3].selectbox("Menú", opciones, index=i, key=k("al", serv, "menu"),
                                              label_visibility="collapsed")
            else:
                cols[3].caption("(sin menús configurados)")
            d["personas"] = cols[4].number_input("Personas", min_value=0, step=1,
                                                 value=entero(d["personas"]),
                                                 key=k("al", serv, "per"),
                                                 label_visibility="collapsed")
        d["valor"] = cols[6].number_input("Valor editable", min_value=0, step=500,
                                          value=entero(d["valor"]), key=k("al", serv, "val"),
                                          label_visibility="collapsed")
        d["domicilio"] = cols[7].number_input("Domicilio", min_value=0, step=1000,
                                              value=entero(d["domicilio"]),
                                              key=k("al", serv, "dom"),
                                              label_visibility="collapsed")
        filas.append((serv, cols))

    res = calcular_alimentacion(cfg, ev)
    for (serv, cols), f in zip(filas, res["filas"]):
        if serv == CAFE:
            cols[4].write(f"{cant(f['unidades'])} estaciones")
        cols[5].write(cop(f["tarifa"]))
        cols[8].write(cop(f["total"]))
    st.markdown(f'<div class="total-box"><small>Total alimentación</small>{cop(res["total"])}</div>',
                unsafe_allow_html=True)


# =============================================================================
#  4. TRANSPORTE
# =============================================================================
def pagina_transporte():
    cfg, ev = CFG(), EV()
    st.title("4. Transporte")
    info("Los vehículos más grandes se llenan primero y el último tipo cubre el remanente. "
         "Ejemplo: 14 personas = 1 aerovan de 12 + 1 taxi.")
    t = ev.setdefault("transporte", {"incluir": False, "propio": 0, "aporta": {}, "valores": {}})
    t["incluir"] = st.toggle("Incluir transporte en la cotización", value=bool(t["incluir"]),
                             key=k("tr", "inc"))
    t["propio"] = st.number_input("Asistentes que llegan por sus propios medios", min_value=0,
                                  step=1, value=entero(t["propio"]), key=k("tr", "propio"))

    anchos = [2.2, 1.2, 1.3, 1.5, 1.2, 1.8, 1.8]
    cabeceras(st.columns(anchos), ["Vehículo", "Capacidad", "Calculados", "Aporta empresa",
                                   "A cotizar", "Valor unitario", "Subtotal"])
    vehiculos = sorted(cfg["vehiculos"], key=lambda v: -num(v["capacidad"]))
    filas = []
    for v in vehiculos:
        cols = st.columns(anchos, vertical_alignment="center")
        cols[0].write(f"**{v['nombre']}**")
        cols[1].write(f"{cant(v['capacidad'])} pers.")
        t["aporta"][v["nombre"]] = cols[3].number_input(
            "Aporta", min_value=0, step=1, value=entero(t["aporta"].get(v["nombre"])),
            key=k("tr", v["nombre"], "apo"), label_visibility="collapsed")
        base = entero(v["valor"])
        actual = entero(t["valores"].get(v["nombre"])) if t["valores"].get(v["nombre"]) is not None else base
        val = cols[5].number_input("Valor", min_value=0, step=5000, value=actual,
                                   key=k("tr", v["nombre"], "val"), label_visibility="collapsed")
        t["valores"][v["nombre"]] = None if val == base else val
        filas.append(cols)

    res = calcular_transporte(cfg, ev)
    for cols, f in zip(filas, res["filas"]):
        cols[2].write(cant(f["calculada"]))
        cols[4].write(cant(f["a_cotizar"]))
        cols[6].write(cop(f["subtotal"] if res["incluir"] else 0))
    st.caption(f"Asistentes que requieren transporte: {cant(res['requieren'])}")
    st.markdown(f'<div class="total-box"><small>Total transporte</small>{cop(res["total"])}</div>',
                unsafe_allow_html=True)


# =============================================================================
#  5. RESUMEN Y EXPORTACIÓN
# =============================================================================
def pagina_resumen():
    cfg, ev = CFG(), EV()
    st.title("5. Resumen y exportación")

    c1, c2, _ = st.columns([1.3, 1, 2])
    ev["iva_incluir"] = c1.toggle("Incluir IVA en la cotización", value=bool(ev["iva_incluir"]),
                                  key=k("iva_inc"))
    ev["iva_pct"] = c2.number_input("IVA (%)", min_value=0.0, max_value=100.0, step=1.0,
                                    value=float(num(ev["iva_pct"])), key=k("iva_pct"),
                                    disabled=not ev["iva_incluir"])
    cot = armar_cotizacion(cfg, ev)
    r = cot["resumen"]

    for a in cot["advertencias"]:
        aviso(a)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Profesionales", cop(r["profesionales"]))
    c2.metric("Otros recursos", cop(r["otros_recursos"]))
    c3.metric("Alimentación", cop(r["alimentacion"]))
    c4.metric("Transporte", cop(r["transporte"]))

    c1, c2 = st.columns(2)
    c1.metric("Subtotal", cop(r["subtotal"]))
    c2.metric(f"IVA ({r['iva_pct']:g}%)" if r["iva_incluido"] else "IVA", cop(r["iva"]),
              help="Activa «Incluir IVA» para sumarlo al total.")
    st.markdown(f'<div class="total-box"><small>TOTAL GENERAL</small>{cop(r["total"])}</div>',
                unsafe_allow_html=True)

    activas = [p for p in cot["pistas"] if p["activa"]]
    if activas:
        st.subheader("Detalle por pista")
        st.dataframe(pd.DataFrame([{"Pista": p["nombre"], "Subtotal": cop(p["subtotal"])}
                                   for p in activas]), hide_index=True, width="stretch")

    st.subheader("Exportar cotización")
    c1, c2 = st.columns(2)
    try:
        c1.download_button("📄 Descargar Word (.docx)", data=generar_word(ev, cot),
                           file_name=nombre_archivo(ev, "docx"),
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           width="stretch")
        c2.download_button("📕 Descargar PDF", data=generar_pdf(ev, cot),
                           file_name=nombre_archivo(ev, "pdf"), mime="application/pdf",
                           width="stretch")
    except Exception as e:
        st.error("No se pudo generar el documento: " + str(e))
    st.caption("El documento lista solo lo que se cotiza: cantidad mayor que cero y valor mayor que cero.")


# =============================================================================
#  CONFIGURACIÓN
# =============================================================================
COLS_RECURSOS = ["nombre", "regla", "personas", "fija", "ref", "divisor", "unidad",
                 "valor", "es_profesional"]


def base_df(clave, filas, columnas, tipos):
    """Copia base de una tabla editable. No cambia mientras se edita (evita filas duplicadas)."""
    snap = st.session_state.snap
    if clave not in snap:
        if filas:
            df = pd.DataFrame(filas)[columnas]
        else:
            df = pd.DataFrame({c: pd.Series(dtype=tipos[c]) for c in columnas})
        for c in columnas:
            if tipos[c] == "float":
                df[c] = pd.to_numeric(df[c], errors="coerce").astype(float)
            elif tipos[c] == "bool":
                df[c] = df[c].fillna(False).astype(bool)
            else:
                df[c] = df[c].fillna("").astype(str)
        snap[clave] = df
    return snap[clave]


def unicos(nombres):
    vistos, salida = {}, []
    for n in nombres:
        if n in vistos:
            vistos[n] += 1
            n = f"{n} ({vistos[n]})"
        else:
            vistos[n] = 1
        salida.append(n)
    return salida


def pestana_parametros(cfg):
    p = cfg["parametros"]
    v = st.session_state.cfg_ver
    c1, c2, c3 = st.columns(3)
    p["tarifa_profesional"] = c1.number_input("Tarifa por profesional ($)", min_value=0, step=1000,
                                              value=entero(p["tarifa_profesional"]), key=f"c{v}_tarifa")
    p["prof_por_asistentes"] = c2.number_input("Asistentes por profesional", min_value=1, step=1,
                                               value=max(1, entero(p["prof_por_asistentes"])),
                                               key=f"c{v}_pxa")
    p["iva_pct"] = c3.number_input("IVA por defecto (%)", min_value=0.0, max_value=100.0, step=1.0,
                                   value=float(num(p["iva_pct"])), key=f"c{v}_iva")
    niveles = st.text_area("Niveles de brigada (uno por línea)",
                           value="\n".join(p["niveles_brigada"]), key=f"c{v}_niv")
    p["niveles_brigada"] = [n.strip() for n in niveles.splitlines() if n.strip()]

    st.markdown("**Estación de café**")
    c1, c2 = st.columns(2)
    cfg["cafe"]["personas_por_estacion"] = c1.number_input(
        "Personas por estación", min_value=1, step=1,
        value=max(1, entero(cfg["cafe"]["personas_por_estacion"])), key=f"c{v}_cafe_p")
    cfg["cafe"]["valor"] = c2.number_input("Valor por estación ($)", min_value=0, step=1000,
                                           value=entero(cfg["cafe"]["valor"]), key=f"c{v}_cafe_v")

    st.markdown("**Horarios para sugerir comidas** (se sugiere si el evento empieza a esa hora "
                "o antes **y** termina a la hora «hasta» o después)")
    for serv in SERVICIOS_ALIMENTACION:
        h = cfg["horarios"].setdefault(serv, {"desde": 8, "hasta": 9})
        c0, c1, c2 = st.columns([1.5, 1, 1])
        c0.write(serv)
        h["desde"] = c1.number_input("Empieza a las (o antes)", min_value=0.0, max_value=24.0,
                                     step=0.5, value=float(num(h["desde"])),
                                     key=f"c{v}_h_{slug(serv)}_d")
        h["hasta"] = c2.number_input("Termina a las (o después)", min_value=0.0, max_value=24.0,
                                     step=0.5, value=float(num(h["hasta"])),
                                     key=f"c{v}_h_{slug(serv)}_h")


def pestana_pistas(cfg):
    v = st.session_state.cfg_ver
    nombres = [p["nombre"] for p in cfg["pistas"]]
    if not nombres:
        st.info("No hay pistas. Agrega una abajo.")
    else:
        sel = st.selectbox("Pista a editar", nombres, key=f"c{v}_pista_sel")
        pista = pista_por_nombre(cfg, sel)
        tipos = {"nombre": "str", "regla": "str", "personas": "float", "fija": "float",
                 "ref": "str", "divisor": "float", "unidad": "str", "valor": "float",
                 "es_profesional": "bool"}
        base = base_df(f"rec_{sel}", pista["recursos"], COLS_RECURSOS, tipos)
        ed = st.data_editor(
            base, key=f"ed{v}_rec_{slug(sel)}", num_rows="dynamic", hide_index=True,
            width="stretch",
            column_config={
                "nombre": st.column_config.TextColumn("Recurso", required=True, width="large"),
                "regla": st.column_config.SelectboxColumn("Regla", options=REGLAS, required=True),
                "personas": st.column_config.NumberColumn("Personas por unidad", min_value=0,
                                                          help="Para la regla «asistentes»"),
                "fija": st.column_config.NumberColumn("Cantidad fija", min_value=0,
                                                      help="Para la regla «fija»"),
                "ref": st.column_config.TextColumn("Recurso de referencia",
                                                   help="Para «proporcional»: nombre exacto de un recurso anterior"),
                "divisor": st.column_config.NumberColumn("Divisor", min_value=0),
                "unidad": st.column_config.TextColumn("Unidad"),
                "valor": st.column_config.NumberColumn("Valor unitario ($)", min_value=0, format="%d"),
                "es_profesional": st.column_config.CheckboxColumn(
                    "Es profesional", help="Usa la tarifa general del profesional"),
            })
        recursos, nombres_r = [], unicos([txt(n).strip() for n in ed["nombre"] if txt(n).strip()])
        i = 0
        for _, f in ed.iterrows():
            if not txt(f["nombre"]).strip():
                continue
            regla = txt(f["regla"]) if txt(f["regla"]) in REGLAS else "asistentes"
            recursos.append({"nombre": nombres_r[i], "regla": regla,
                             "personas": num(f["personas"]), "fija": num(f["fija"]),
                             "ref": txt(f["ref"]).strip(), "divisor": num(f["divisor"]),
                             "unidad": txt(f["unidad"]).strip() or "unidad",
                             "valor": num(f["valor"]), "es_profesional": bool(f["es_profesional"])})
            i += 1
        pista["recursos"] = recursos
        with st.expander("¿Qué significa cada regla?"):
            for r, ayuda in REGLAS_AYUDA.items():
                st.markdown(f"- **{r}**: {ayuda}")
        st.caption("Para agregar una fila usa la última fila vacía de la tabla; para borrar, "
                   "selecciona la fila y presiona Supr.")

        with st.expander("Eliminar esta pista"):
            if st.checkbox(f"Confirmo que quiero eliminar «{sel}»", key=f"c{v}_del_{slug(sel)}"):
                if st.button("Eliminar pista", key=f"c{v}_delbtn_{slug(sel)}"):
                    cfg["pistas"] = [p for p in cfg["pistas"] if p["nombre"] != sel]
                    st.session_state.snap.pop(f"rec_{sel}", None)
                    st.rerun()

    st.markdown("**Agregar una pista nueva**")
    c1, c2 = st.columns([3, 1])
    nueva = c1.text_input("Nombre de la pista nueva", key=f"c{v}_nueva_pista")
    if c2.button("Agregar pista", key=f"c{v}_add_pista"):
        n = nueva.strip()
        if not n:
            st.warning("Escribe un nombre.")
        elif n in [p["nombre"] for p in cfg["pistas"]]:
            st.warning("Ya existe una pista con ese nombre.")
        else:
            cfg["pistas"].append({"nombre": n, "recursos": [
                {"nombre": "Profesionales", "regla": "profesionales", "personas": 0, "fija": 0,
                 "ref": "", "divisor": 0, "unidad": "profesional", "valor": 0,
                 "es_profesional": True}]})
            st.rerun()
    st.caption("Para cambiar el nombre de una pista, créala nueva y elimina la anterior.")


def pestana_menus(cfg):
    v = st.session_state.cfg_ver
    serv = st.selectbox("Servicio", SERVICIOS_ALIMENTACION, key=f"c{v}_menu_serv")
    tipos = {"nombre": "str", "valor": "float"}
    base = base_df(f"menu_{serv}", cfg["menus"].get(serv, []), ["nombre", "valor"], tipos)
    ed = st.data_editor(base, key=f"ed{v}_menu_{slug(serv)}", num_rows="dynamic", hide_index=True,
                        width="stretch",
                        column_config={
                            "nombre": st.column_config.TextColumn("Menú", required=True),
                            "valor": st.column_config.NumberColumn("Valor por persona ($)",
                                                                   min_value=0, format="%d")})
    filas = [(txt(a).strip(), b) for a, b in zip(ed["nombre"], ed["valor"]) if txt(a).strip()]
    nombres = unicos([f[0] for f in filas])
    cfg["menus"][serv] = [{"nombre": n, "valor": num(val)}
                          for n, (_, val) in zip(nombres, filas)]


def pestana_vehiculos(cfg):
    v = st.session_state.cfg_ver
    tipos = {"nombre": "str", "capacidad": "float", "valor": "float"}
    base = base_df("vehiculos", cfg["vehiculos"], ["nombre", "capacidad", "valor"], tipos)
    ed = st.data_editor(base, key=f"ed{v}_vehiculos", num_rows="dynamic", hide_index=True,
                        width="stretch",
                        column_config={
                            "nombre": st.column_config.TextColumn("Vehículo", required=True),
                            "capacidad": st.column_config.NumberColumn("Capacidad (personas)",
                                                                       min_value=1),
                            "valor": st.column_config.NumberColumn("Valor por vehículo ($)",
                                                                   min_value=0, format="%d")})
    filas = [(txt(a).strip(), b, c) for a, b, c in zip(ed["nombre"], ed["capacidad"], ed["valor"])
             if txt(a).strip()]
    nombres = unicos([f[0] for f in filas])
    cfg["vehiculos"] = [{"nombre": n, "capacidad": max(1.0, num(cap, 1)), "valor": num(val)}
                        for n, (_, cap, val) in zip(nombres, filas)]
    st.caption("Se llenan primero los vehículos de mayor capacidad; el de menor capacidad "
               "cubre el remanente.")


def pestana_archivo(cfg):
    st.markdown("**Guardar la configuración**")
    data = {"tipo": "kumara_configuracion", "version": 1,
            "guardado": datetime.now().isoformat(timespec="seconds"), "config": cfg}
    st.download_button("💾 Descargar configuración (.json)",
                       data=json.dumps(data, ensure_ascii=False, indent=2),
                       file_name="config_kumara.json", mime="application/json")
    info("Para que esta configuración sea la <b>inicial para todos</b>, sube ese archivo a tu "
         "repositorio de GitHub con el nombre <code>config_kumara.json</code> (junto a app.py).")

    st.markdown("**Cargar una configuración**")
    up = st.file_uploader("Archivo de configuración (.json)", type="json", key="up_cfg_file")
    if up is not None:
        firma = getattr(up, "file_id", None) or f"{up.name}-{up.size}"
        if st.session_state.get("firma_cfg") != firma:
            st.session_state.firma_cfg = firma
            try:
                datos = json.loads(up.getvalue().decode("utf-8"))
                if datos.get("tipo") != "kumara_configuracion":
                    raise ValueError("el archivo no es una configuración de Kumara")
                st.session_state.config = fusionar(config_defecto(), datos["config"])
                st.session_state.cfg_ver += 1
                st.session_state.snap = {}
                st.session_state.mensaje = "Configuración cargada."
                st.rerun()
            except Exception as e:
                st.error("No se pudo cargar: " + str(e))

    st.markdown("**Restaurar valores originales**")
    if st.checkbox("Confirmo que quiero volver a los valores del Excel original",
                   key=f"c{st.session_state.cfg_ver}_conf_rest"):
        if st.button("Restaurar configuración"):
            st.session_state.config = config_defecto()
            st.session_state.cfg_ver += 1
            st.session_state.snap = {}
            st.session_state.mensaje = "Configuración restaurada."
            st.rerun()


def pagina_configuracion():
    cfg = CFG()
    st.title("Configuración")
    info("Aquí está todo lo que <b>no cambia</b> de un evento a otro. Los cambios se aplican "
         "al instante en las cotizaciones. Descarga la configuración para no perderla.")
    t1, t2, t3, t4, t5 = st.tabs(["Parámetros", "Pistas y recursos", "Menús", "Vehículos",
                                  "Guardar / cargar"])
    with t1:
        pestana_parametros(cfg)
    with t2:
        pestana_pistas(cfg)
    with t3:
        pestana_menus(cfg)
    with t4:
        pestana_vehiculos(cfg)
    with t5:
        pestana_archivo(cfg)


# =============================================================================
#  BARRA LATERAL Y ENRUTADOR
# =============================================================================
PAGINAS = ["Inicio / Guía", "1. Datos del evento", "2. Pistas", "3. Alimentación",
           "4. Transporte", "5. Resumen y exportación", "Configuración"]


def pedir_nueva():
    st.session_state.pedir_nueva = True


def cancelar_nueva():
    st.session_state.pedir_nueva = False


def confirmar_nueva():
    ss = st.session_state
    ss.evento = evento_defecto(ss.config)
    ss.ver += 1
    ss.pedir_nueva = False
    ss.mensaje = "Cotización nueva lista."


with st.sidebar:
    if os.path.exists(LOGO):
        st.image(LOGO, width=150)
    st.markdown("## 🧯 Kumara")
    st.caption("Cotizador de eventos de emergencias")
    pagina = st.radio("Navegación", PAGINAS, label_visibility="collapsed")
    st.divider()
    caja_total = st.empty()        # se llena después de ejecutar la página
    st.divider()
    st.markdown("**Cotización**")
    if st.session_state.mensaje:
        st.success(st.session_state.mensaje)
        st.session_state.mensaje = ""

    archivo = st.file_uploader("Cargar cotización (.json)", type="json", key="up_ev")
    con_cfg = st.checkbox("Usar también la configuración guardada en el archivo", key="up_con_cfg")
    if archivo is not None:
        firma = getattr(archivo, "file_id", None) or f"{archivo.name}-{archivo.size}"
        if st.session_state.get("firma_ev") != firma:
            st.session_state.firma_ev = firma
            try:
                datos = json.loads(archivo.getvalue().decode("utf-8"))
                if datos.get("tipo") != "kumara_cotizacion":
                    raise ValueError("el archivo no es una cotización de Kumara")
                if con_cfg and "config" in datos:
                    st.session_state.config = fusionar(config_defecto(), datos["config"])
                    st.session_state.cfg_ver += 1
                    st.session_state.snap = {}
                st.session_state.evento = fusionar(evento_defecto(st.session_state.config),
                                                   datos["evento"])
                st.session_state.ver += 1
                st.session_state.mensaje = "Cotización cargada correctamente."
                st.rerun()
            except Exception as e:
                st.error("No se pudo cargar: " + str(e))
    caja_guardar = st.container()   # se llena después de ejecutar la página

    if st.session_state.pedir_nueva:
        st.warning("¿Borrar los datos del evento? La configuración se conserva.")
        c1, c2 = st.columns(2)
        c1.button("Sí, borrar", on_click=confirmar_nueva)
        c2.button("Cancelar", on_click=cancelar_nueva)
    else:
        st.button("🆕 Nueva cotización", on_click=pedir_nueva)

# --- página activa ---
{"Inicio / Guía": pagina_inicio,
 "1. Datos del evento": pagina_evento,
 "2. Pistas": pagina_pistas,
 "3. Alimentación": pagina_alimentacion,
 "4. Transporte": pagina_transporte,
 "5. Resumen y exportación": pagina_resumen,
 "Configuración": pagina_configuracion}[pagina]()

# --- elementos de la barra lateral que dependen de lo que acaba de ejecutarse ---
cot_actual = armar_cotizacion(CFG(), EV())
caja_total.markdown(
    f'<div class="total-box" style="font-size:1.15em;color:#fff;">'
    f'<small>TOTAL GENERAL</small>{cop(cot_actual["resumen"]["total"])}</div>',
    unsafe_allow_html=True)
with caja_guardar:
    paquete = {"tipo": "kumara_cotizacion", "version": 1,
               "guardado": datetime.now().isoformat(timespec="seconds"),
               "evento": EV(), "config": CFG()}
    st.download_button("💾 Guardar cotización (.json)",
                       data=json.dumps(paquete, ensure_ascii=False, indent=2),
                       file_name=nombre_archivo(EV(), "json"), mime="application/json")

st.markdown('<div class="pie">Kumara · Cotizador de eventos de emergencias</div>',
            unsafe_allow_html=True)
