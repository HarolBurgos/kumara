"""
modelos.py  -  Kumara | Cotizador de eventos de emergencias
Configuración por defecto (editable desde la app) y todos los cálculos.
La interfaz (app.py) y la exportación (exportar.py) solo consumen estas funciones.
"""
import copy
import math
import re

SI = "Sí"
NO = "No"

# Reglas con las que se calcula la "cantidad requerida" de cada recurso
REGLAS = ["asistentes", "fija", "profesionales", "proporcional"]
REGLAS_AYUDA = {
    "asistentes": "asistentes ÷ personas por unidad (redondea hacia arriba)",
    "fija": "cantidad fija, sin importar los asistentes",
    "profesionales": "asistentes ÷ 'profesionales por asistentes' (parámetro general)",
    "proporcional": "cantidad a cotizar de otro recurso ÷ divisor (el otro recurso debe estar antes)",
}

SERVICIOS_ALIMENTACION = ["Desayuno", "Refrigerio AM", "Almuerzo",
                          "Refrigerio PM", "Cena"]
CAFE = "Estación de café"


# =============================================================================
#  CONFIGURACIÓN POR DEFECTO  (tomada del Excel Cotizador_Kumara_version_final)
# =============================================================================
def _r(nombre, regla="asistentes", personas=0, fija=0, ref="", divisor=0,
       unidad="unidad", valor=0, prof=False):
    return {"nombre": nombre, "regla": regla, "personas": personas,
            "fija": fija, "ref": ref, "divisor": divisor,
            "unidad": unidad, "valor": valor, "es_profesional": prof}


def _profesionales(regla="profesionales", fija=0):
    return _r("Profesionales", regla=regla, fija=fija,
              unidad="profesional", prof=True)


def config_defecto():
    return {
        "parametros": {
            "tarifa_profesional": 52000,
            "prof_por_asistentes": 35,
            "niveles_brigada": ["Básico", "Intermedio", "Avanzado"],
            "iva_pct": 19.0,
        },
        "pistas": [
            {"nombre": "Calentamiento inicial", "recursos": [
                _profesionales(regla="fija", fija=1),
                _r("Parlante / equipo de sonido", personas=150, valor=20000),
                _r("Micrófono", personas=150, valor=20000),
                _r("Conos de señalización", personas=20, valor=20000),
                _r("Bandas elásticas", personas=20, valor=20000),
                _r("Colchonetas", personas=20, valor=20000),
            ]},
            {"nombre": "Incendios", "recursos": [
                _profesionales(),
                _r("Extintores", personas=5, valor=20000),
                _r("Gasolina", personas=50, unidad="galón", valor=14000),
                _r("ACPM", regla="proporcional", ref="Extintores", divisor=2,
                   unidad="galón", valor=9000),
                _r("Canecas", regla="proporcional", ref="Gasolina", divisor=2,
                   valor=20000),
                _r("Tanque de gas propano", personas=120, valor=35000),
            ]},
            {"nombre": "Primeros auxilios", "recursos": [
                _profesionales(),
                _r("Alquiler camillas", personas=6, valor=20000),
                _r("Alquiler botiquines", personas=6, valor=20000),
                _r("Alquiler equipos RCP", personas=6, valor=20000),
                _r("Botellas plásticas", personas=6),
                _r("Anilinas"),
                _r("Torniquetes"),
            ]},
            {"nombre": "Psicología de la emergencia", "recursos": [
                _profesionales(),
                _r("Material didáctico", personas=20),
                _r("Tarjetas de situaciones de emergencia", personas=20),
                _r("Hojas y marcadores", personas=20),
            ]},
            {"nombre": "Manejo de químicos", "recursos": [
                _profesionales(),
                _r("Kit de derrames", personas=20, valor=20000),
                _r("Guantes de protección química"),
                _r("Gafas de seguridad"),
                _r("Material absorbente"),
                _r("Bolsas para residuos contaminados"),
            ]},
            {"nombre": "Prueba de confianza", "recursos": [
                _profesionales(),
                _r("Material para arrastre (estacas, polisombra, clavos, etc.)",
                   regla="fija", fija=8, valor=20000),
                _r("Antifaces", personas=1, valor=1000),
                _r("Cuerdas de seguridad", personas=1, valor=12000),
            ]},
        ],
        "menus": {
            "Desayuno": [
                {"nombre": "Continental", "valor": 12000},
                {"nombre": "Americano", "valor": 16000},
                {"nombre": "Típico regional", "valor": 18000},
                {"nombre": "Saludable", "valor": 15000},
                {"nombre": "Sándwich y café", "valor": 13000},
            ],
            "Refrigerio AM": [
                {"nombre": "Fruta y bebida", "valor": 7000},
                {"nombre": "Sándwich y jugo", "valor": 23000},
                {"nombre": "Yogur y cereal", "valor": 9000},
                {"nombre": "Empanada y bebida", "valor": 8000},
                {"nombre": "Galletas y café", "valor": 6500},
            ],
            "Almuerzo": [
                {"nombre": "Corriente", "valor": 18000},
                {"nombre": "Ejecutivo", "valor": 25000},
                {"nombre": "Vegetariano", "valor": 22000},
                {"nombre": "Pollo", "valor": 24000},
                {"nombre": "Carne", "valor": 28000},
            ],
            "Refrigerio PM": [
                {"nombre": "Fruta y bebida", "valor": 7000},
                {"nombre": "Sándwich y jugo", "valor": 11000},
                {"nombre": "Yogur y cereal", "valor": 9000},
                {"nombre": "Empanada y bebida", "valor": 8000},
                {"nombre": "Galletas y café", "valor": 6500},
            ],
            "Cena": [
                {"nombre": "Cena ligera", "valor": 17000},
                {"nombre": "Menú ejecutivo", "valor": 25000},
                {"nombre": "Vegetariana", "valor": 22000},
                {"nombre": "Pollo", "valor": 24000},
                {"nombre": "Carne", "valor": 28000},
            ],
        },
        "cafe": {"personas_por_estacion": 50, "valor": 50000},
        # servicio sugerido si hora_inicio <= desde  y  hora_fin >= hasta
        "horarios": {
            "Desayuno": {"desde": 8, "hasta": 9},
            "Refrigerio AM": {"desde": 10, "hasta": 11},
            "Almuerzo": {"desde": 12, "hasta": 14},
            "Refrigerio PM": {"desde": 15, "hasta": 16},
            "Cena": {"desde": 18, "hasta": 19},
        },
        "vehiculos": [
            {"nombre": "Aerovan", "capacidad": 12, "valor": 200000},
            {"nombre": "Taxi", "capacidad": 4, "valor": 40000},
        ],
    }


def evento_defecto(config=None):
    config = config or config_defecto()
    return {
        "empresa": "", "contacto": "", "ciudad": "", "fecha": "",
        "lugar": "", "direccion": "",
        "asistentes": 0, "h_inicio": 8.0, "h_fin": 17.0,
        "nivel": config["parametros"]["niveles_brigada"][0],
        "pistas": {},
        "alimentacion": {},
        "transporte": {"incluir": False, "propio": 0, "aporta": {}, "valores": {}},
        "iva_incluir": False,
        "iva_pct": config["parametros"]["iva_pct"],
    }


def fusionar(base, nuevo):
    """Mezcla 'nuevo' sobre 'base' (para cargar archivos viejos sin perder claves)."""
    salida = copy.deepcopy(base)
    for k, v in (nuevo or {}).items():
        if isinstance(v, dict) and isinstance(salida.get(k), dict):
            salida[k] = fusionar(salida[k], v)
        else:
            salida[k] = copy.deepcopy(v)
    return salida


# =============================================================================
#  UTILIDADES
# =============================================================================
def num(v, defecto=0.0):
    try:
        if v is None or v == "":
            return defecto
        f = float(v)
        return defecto if f != f else f      # NaN -> defecto
    except (TypeError, ValueError):
        return defecto


def txt(v, defecto=""):
    """Texto seguro (None / NaN -> defecto)."""
    if v is None or (isinstance(v, float) and v != v):
        return defecto
    return str(v)


def cop(v):
    """Formato pesos colombianos: $ 1.234.567"""
    return "$ " + f"{round(num(v)):,}".replace(",", ".")


def cant(v):
    """Cantidad sin decimales innecesarios."""
    v = num(v)
    return str(int(v)) if v == int(v) else f"{v:.2f}"


def slug(texto):
    return re.sub(r"[^a-z0-9]+", "_", str(texto).lower()).strip("_")


def pista_por_nombre(config, nombre):
    return next((p for p in config["pistas"] if p["nombre"] == nombre), None)


def ov_pista(evento, nombre):
    """Datos del evento para una pista (se crean si no existen)."""
    return evento["pistas"].setdefault(
        nombre, {"activa": False, "recursos": {}, "extras": []})


def ov_recurso(evento, pista, recurso):
    return ov_pista(evento, pista)["recursos"].setdefault(
        recurso, {"aporta": 0, "incluir": True, "obs": "", "valor": None})


# =============================================================================
#  PISTAS
# =============================================================================
def cantidad_requerida(rec, asistentes, params, cache):
    regla = rec.get("regla", "asistentes")
    if regla == "fija":
        return num(rec.get("fija"))
    if regla == "profesionales":
        if asistentes <= 0:
            return 0
        return math.ceil(asistentes / max(1.0, num(params["prof_por_asistentes"], 1)))
    if regla == "proporcional":
        base = cache.get(rec.get("ref", ""), 0)
        div = num(rec.get("divisor"))
        return math.ceil(base / div) if div > 0 else 0
    # "asistentes"
    p = num(rec.get("personas"))
    if asistentes <= 0 or p <= 0:
        return 0
    return math.ceil(asistentes / p)


def valor_base(rec, params):
    """Valor unitario del catálogo (profesionales usan la tarifa general)."""
    if rec.get("es_profesional"):
        return num(params["tarifa_profesional"])
    return num(rec.get("valor"))


def calcular_linea(rec, o, asistentes, params, cache, activa=True):
    """Calcula una fila de recurso. 'o' son los datos del evento para esa fila."""
    aporta = num(o.get("aporta"))
    incluir = bool(o.get("incluir", True))
    valor = num(o["valor"]) if o.get("valor") is not None else valor_base(rec, params)
    req = cantidad_requerida(rec, asistentes, params, cache)
    a_cot = max(0, req - aporta) if incluir else 0
    return {
        "nombre": rec["nombre"], "unidad": rec.get("unidad", ""),
        "personas": rec.get("personas", 0), "regla": rec.get("regla"),
        "requerida": req, "aporta": aporta, "a_cotizar": a_cot,
        "valor": valor, "subtotal": a_cot * valor if (activa and incluir) else 0,
        "incluir": incluir, "obs": o.get("obs", ""),
        "es_profesional": bool(rec.get("es_profesional")), "extra": False,
    }


def calcular_pista(config, evento, nombre_pista):
    pista = pista_por_nombre(config, nombre_pista)
    datos = evento["pistas"].get(nombre_pista, {})
    activa = bool(datos.get("activa", False))
    params = config["parametros"]
    asistentes = num(evento.get("asistentes"))
    lineas, cache = [], {}

    for rec in pista["recursos"]:
        o = datos.get("recursos", {}).get(rec["nombre"], {})
        linea = calcular_linea(rec, o, asistentes, params, cache, activa)
        cache[rec["nombre"]] = linea["a_cotizar"]
        lineas.append(linea)

    for ex in datos.get("extras", []):
        c, v = num(ex.get("cantidad")), num(ex.get("valor"))
        lineas.append({
            "nombre": ex.get("nombre", "Recurso adicional"), "unidad": "",
            "personas": 0, "regla": "extra", "requerida": c, "aporta": 0,
            "a_cotizar": c, "valor": v, "subtotal": c * v if activa else 0,
            "incluir": True, "obs": ex.get("obs", ""),
            "es_profesional": False, "extra": True,
        })

    subtotal = sum(l["subtotal"] for l in lineas) if activa else 0
    return {"nombre": nombre_pista, "activa": activa, "lineas": lineas,
            "subtotal": subtotal}


# =============================================================================
#  ALIMENTACIÓN
# =============================================================================
def sugerido(config, servicio, h_inicio, h_fin):
    h = config["horarios"].get(servicio)
    if not h:
        return ""
    return SI if (num(h_inicio) <= num(h["desde"]) and num(h_fin) >= num(h["hasta"])) else NO


def calcular_alimentacion(config, evento):
    asistentes = num(evento.get("asistentes"))
    filas = []
    for serv in SERVICIOS_ALIMENTACION + [CAFE]:
        d = evento.get("alimentacion", {}).get(serv, {})
        incluir = bool(d.get("incluir", False))
        edit = num(d.get("valor"))
        dom = num(d.get("domicilio"))
        if serv == CAFE:
            cafe = config["cafe"]
            por = max(1.0, num(cafe["personas_por_estacion"], 1))
            unidades = math.ceil(asistentes / por) if asistentes > 0 else 0
            tarifa = num(cafe["valor"])
            menu = "Estación de café"
        else:
            menu = d.get("menu") or ""
            lista = config["menus"].get(serv, [])
            if not menu and lista:
                menu = lista[0]["nombre"]
            tarifa = next((num(m["valor"]) for m in lista if m["nombre"] == menu), 0.0)
            unidades = num(d.get("personas")) or asistentes
        valor = edit if edit > 0 else tarifa
        total = unidades * valor + dom if incluir else 0
        filas.append({
            "servicio": serv, "menu": menu, "incluir": incluir,
            "sugerido": "" if serv == CAFE else sugerido(
                config, serv, evento.get("h_inicio"), evento.get("h_fin")),
            "unidades": unidades, "tarifa": tarifa, "valor": valor,
            "domicilio": dom, "total": total,
        })
    return {"filas": filas, "total": sum(f["total"] for f in filas)}


# =============================================================================
#  TRANSPORTE
# =============================================================================
def calcular_transporte(config, evento):
    t = evento.get("transporte", {})
    incluir = bool(t.get("incluir", False))
    asistentes = num(evento.get("asistentes"))
    propio = num(t.get("propio"))
    requieren = max(0.0, asistentes - propio)
    vehiculos = sorted(config["vehiculos"], key=lambda v: -num(v["capacidad"]))
    filas, restante = [], requieren
    for i, v in enumerate(vehiculos):
        capac = max(1.0, num(v["capacidad"], 1))
        if not incluir:
            n = 0
        elif i < len(vehiculos) - 1:
            n = int(restante // capac)
            restante -= n * capac
        else:
            n = math.ceil(restante / capac) if restante > 0 else 0
        aporta = num(t.get("aporta", {}).get(v["nombre"]))
        a_cot = max(0, n - aporta)
        ov = t.get("valores", {}).get(v["nombre"])
        valor = num(ov) if ov is not None else num(v["valor"])
        filas.append({
            "vehiculo": v["nombre"], "capacidad": capac, "calculada": n,
            "aporta": aporta, "a_cotizar": a_cot, "valor": valor,
            "subtotal": a_cot * valor,
        })
    return {"incluir": incluir, "requieren": requieren, "filas": filas,
            "total": sum(f["subtotal"] for f in filas) if incluir else 0}


# =============================================================================
#  RESUMEN GENERAL
# =============================================================================
def armar_cotizacion(config, evento):
    pistas = [calcular_pista(config, evento, p["nombre"]) for p in config["pistas"]]
    activas = [p for p in pistas if p["activa"]]
    profesionales = sum(l["subtotal"] for p in activas for l in p["lineas"]
                        if l["es_profesional"])
    total_pistas = sum(p["subtotal"] for p in activas)
    alim = calcular_alimentacion(config, evento)
    trans = calcular_transporte(config, evento)
    subtotal = total_pistas + alim["total"] + trans["total"]
    iva = subtotal * num(evento.get("iva_pct")) / 100.0 if evento.get("iva_incluir") else 0.0

    advertencias = []
    if num(evento.get("asistentes")) <= 0:
        advertencias.append("Falta el número de asistentes: las cantidades no se pueden calcular.")
    if num(evento.get("h_fin")) <= num(evento.get("h_inicio")):
        advertencias.append("La hora de fin debe ser mayor que la hora de inicio.")
    for p in activas:
        for l in p["lineas"]:
            if l["incluir"] and l["a_cotizar"] > 0 and l["valor"] <= 0:
                advertencias.append(
                    f"{p['nombre']}: «{l['nombre']}» tiene cantidad pero no tiene valor "
                    f"unitario (no suma al total).")

    return {
        "pistas": pistas,
        "alimentacion": alim,
        "transporte": trans,
        "resumen": {
            "profesionales": profesionales,
            "otros_recursos": total_pistas - profesionales,
            "alimentacion": alim["total"],
            "transporte": trans["total"],
            "subtotal": subtotal,
            "iva_incluido": bool(evento.get("iva_incluir")),
            "iva_pct": num(evento.get("iva_pct")),
            "iva": iva,
            "total": subtotal + iva,
        },
        "duracion": max(0.0, num(evento.get("h_fin")) - num(evento.get("h_inicio"))),
        "advertencias": advertencias,
    }
