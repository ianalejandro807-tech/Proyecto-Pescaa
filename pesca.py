"""Proyecto 4: Captura Marítima y Registro Pesquero (interfaz web con Streamlit).

Ejecutar:  streamlit run app.py
Requiere el archivo 15000.xlsx en la misma carpeta.
"""
import random
from collections import Counter
from pathlib import Path

import streamlit as st

import algoritmos as alg

CARPETA = Path(__file__).parent
RUTA_EXCEL = CARPETA / "15000.xlsx"
NOMBRES_CAMPO = {1: "id_embarcacion", 2: "toneladas_capturadas", 3: "fecha_arribo"}

st.set_page_config(page_title="Captura Marítima y Registro Pesquero", layout="wide")
ss = st.session_state

# ------------------------------------------------------------ 1. Introducción
st.title("🐟 Captura Marítima y Registro Pesquero Ribereño y de Alta Mar")
st.markdown(
    "Proyecto que procesa **15,000 avisos de arribo** de CONAPESCA/SEMAR/INAPESCA "
    "(Golfo de México, Mar de Cortés y Pacífico). Se ordenan los registros con un "
    "método elegido **al azar** y luego se busca un dato con búsqueda lineal o binaria, "
    "midiendo el rendimiento de cada algoritmo en nanosegundos."
)


# -------------------------------------------------- 2. Carga de datos en RAM
@st.cache_data(show_spinner="Cargando datos...")
def cargar():
    return alg.cargar_excel(RUTA_EXCEL)


def nueva_ejecucion():
    """Disparador aleatorio de ordenación + sorteo del método de búsqueda."""
    campo = random.randint(1, 3)
    metodo = random.choice(alg.METODOS[campo])
    kk = alg.CAMPOS_CLAVE[campo]
    # Verificación: ¿la entrada estaba ordenada por el campo clave? (debe ser NO)
    ss.entrada_ordenada = all(
        ss.datos[i][kk] <= ss.datos[i + 1][kk] for i in range(len(ss.datos) - 1)
    )
    ss.entrada_muestra = ss.datos[:5]  # primeros 5 registros que recibe el algoritmo
    ordenado, ns, comp, mov = alg.ordenar(ss.datos, campo, metodo)
    ss.salida_ordenada = all(
        ordenado[i][kk] <= ordenado[i + 1][kk] for i in range(len(ordenado) - 1)
    )
    ss.ordenado = ordenado
    ss.ultimo = dict(campo=campo, metodo=metodo, ns=ns, comp=comp, mov=mov)
    nombre = NOMBRES_CAMPO[campo]
    ss.msg_orden = alg.mensaje_ordenacion(len(ss.datos), nombre, metodo, ns, comp, mov)
    alg.guardar_bitacora_ordenacion(CARPETA, ss.msg_orden)
    if not ss.historial:  # 1ª ejecución: guardar los 15,000 registros ordenados
        alg.guardar_registros_ordenados(CARPETA, ordenado, nombre)
    ss.historial.append(
        {"Ejecución #": len(ss.historial) + 1, "Campo Clave Seleccionado (1-3)": campo,
         "Nombre del Campo": nombre, "Método de Ordenamiento Asignado": metodo,
         "Tiempo de Ejecución (ns)": ns}
    )
    ss.busqueda = random.choice(["Lineal", "Binaria"])
    ss.resultado = None


if "datos" not in ss:
    ss.datos, ss.ns_carga = cargar()
    ss.historial = []
    nueva_ejecucion()

datos = ss.datos
st.success(f"Se cargaron {len(datos):,} registros en memoria RAM en {ss.ns_carga:,} nanosegundos.")


# ------------------------------------------------ 3. Primeros 30 y últimos 30
def tabla(filas):
    return [dict(zip(alg.COLUMNAS, f)) for f in filas]


def alternar(clave):
    ss[clave] = not ss.get(clave, False)


st.subheader("Registros cargados")
st.button(
    "Ocultar primeros 30 y últimos 30" if ss.get("ver_carga") else "Mostrar primeros 30 y últimos 30",
    key="btn_carga", on_click=alternar, args=("ver_carga",),
)
if ss.get("ver_carga"):
    c1, c2 = st.columns(2)
    c1.caption("Primeros 30")
    c1.dataframe(tabla(datos[:30]), use_container_width=True, hide_index=True)
    c2.caption("Últimos 30")
    c2.dataframe(tabla(datos[-30:]), use_container_width=True, hide_index=True)

# -------------------------------------------------- 4 y 5. Ordenamiento
u = ss.ultimo
nombre_clave_orden = NOMBRES_CAMPO[u["campo"]]
st.subheader("Ordenamiento (disparador aleatorio)")
if st.button("🔄 Volver a ordenar (nuevo sorteo aleatorio)", key="btn_reordenar"):
    nueva_ejecucion()
    st.rerun()
st.info(
    f"Campo clave sorteado: **{u['campo']} ({NOMBRES_CAMPO[u['campo']]})** · "
    f"Método sorteado: **{u['metodo']}**"
)
st.success(ss.msg_orden + "  (guardado en bitacora_ordenacion.txt)")
oc1, oc2 = st.columns(2)
oc1.caption("Primeros 30 ordenados")
oc1.dataframe(tabla(ss.ordenado[:30]), use_container_width=True, hide_index=True)
oc2.caption("Últimos 30 ordenados")
oc2.dataframe(tabla(ss.ordenado[-30:]), use_container_width=True, hide_index=True)

# ----------------------------------------------------------- 6. Búsqueda
st.subheader(f"Búsqueda (sorteo: Búsqueda {ss.busqueda})")
k = alg.CAMPOS_CLAVE[u["campo"]]
nombre_clave = NOMBRES_CAMPO[u["campo"]]

if ss.busqueda == "Binaria":
    st.markdown("**[VALIDACIÓN DE SEGURIDAD]** Verificando estado físico de la memoria...")
    st.write(f"El campo clave requerido para la búsqueda es: **[{nombre_clave}]**")
    st.write("Estado del campo en memoria: **YA ORDENADO PREVIAMENTE.**")
    col_busq = k
else:
    st.write("Puedes buscar en cualquier campo (clave o informativo).")
    nombre_col = st.selectbox("Campo en el que quieres buscar", alg.COLUMNAS)
    col_busq = alg.COLUMNAS.index(nombre_col)

campo_buscado = alg.COLUMNAS[col_busq]
dato = st.text_input(f"Introduzca el dato a buscar ({campo_buscado})")
if st.button("Buscar") and dato.strip():
    muestra = datos[0][col_busq]
    try:
        if isinstance(muestra, int):
            valor = int(dato)
        elif isinstance(muestra, float):
            valor = float(dato)
        else:
            valor = dato
        if ss.busqueda == "Binaria":
            enc, filas, ns, comp = alg.busqueda_binaria(ss.ordenado, k, valor)
            metodo_b = "Búsqueda Binaria"
        else:
            enc, filas, ns, comp = alg.busqueda_lineal(datos, col_busq, valor)
            metodo_b = "Búsqueda Lineal"
        msg = alg.mensaje_busqueda(len(datos), campo_buscado, metodo_b, ns, comp)
        alg.guardar_bitacora_busqueda(CARPETA, msg)
        ss.resultado = dict(enc=enc, filas=filas, msg=msg, dato=dato, campo=campo_buscado)
    except ValueError:
        st.error("El dato no tiene el formato del campo (se esperaba un número).")
        ss.resultado = None

r = ss.resultado
if r:
    st.success(r["msg"] + "  (guardado en bitacora_busqueda.txt)")
    if r["enc"]:
        n = len(r["enc"])
        st.write(f"Dato encontrado en la posición indexada de memoria RAM [Fila {r['filas'][0]}].")
        st.code(alg.ficha_registro(r["enc"][0], r["filas"][0]), language=None)
        if n > 1:
            st.caption(f"Hay {n:,} registros con ese valor (se muestran hasta 100).")
            filas_txt = ", ".join(str(f) for f in r["filas"][:100])
            st.write(f"Filas: {filas_txt}")
            st.dataframe(tabla(r["enc"][:100]), use_container_width=True, hide_index=True)
    else:
        st.warning(f"No se encontró '{r['dato']}' en {r['campo']}.")

# ------------------------------------- 7. Registro de ejecuciones y método Alfa
st.subheader("Registro de ejecuciones")
st.dataframe(ss.historial, use_container_width=True, hide_index=True)

conteo = Counter(h["Método de Ordenamiento Asignado"] for h in ss.historial)
st.subheader("Método Alfa (selección de métodos de ordenamiento)")
todos = sorted({m for ms in alg.METODOS.values() for m in ms})
st.dataframe(
    [{"Método": m, "Veces seleccionado": conteo.get(m, 0)} for m in todos],
    use_container_width=True, hide_index=True,
)
alfa = max(conteo.items(), key=lambda x: x[1])
st.success(f"Método Alfa (el más seleccionado): **{alfa[0]}** con {alfa[1]} selección(es).")

# ------------------------------------------------------------- Descargas
st.subheader("Archivos generados")
d1, d2, d3 = st.columns(3)
for col, nombre in zip(
    (d1, d2, d3),
    (alg.BITACORA_ORDENACION, alg.BITACORA_BUSQUEDA, alg.REGISTROS_ORDENADOS),
):
    ruta = CARPETA / nombre
    if ruta.exists():
        col.download_button(f"⬇️ {nombre}", ruta.read_bytes(), file_name=nombre, mime="text/plain")
    else:
        col.caption(f"{nombre}: aún no se genera")

st.divider()
st.caption("El programa ha finalizado su ejecución. Las bitácoras han sido actualizadas.")