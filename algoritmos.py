"""Algoritmos de ordenamiento y búsqueda con conteo de comparaciones y movimientos.

Cada algoritmo ordena IN-PLACE (o devuelve la lista ordenada) una lista de
registros (listas/tuplas) usando la columna `k` como llave, y devuelve
(registros_ordenados, comparaciones, movimientos).
"""
import time
from pathlib import Path

import openpyxl

COLUMNAS = [
    "id_embarcacion", "toneladas_capturadas", "fecha_arribo", "especie_maritima",
    "puerto_desembarque", "estado", "nombre_permisionario", "tipo_arte_pesca",
    "dias_alta_mar", "precio_kilo_venta",
]
CAMPOS_CLAVE = {1: 0, 2: 1, 3: 2}  # número de campo clave -> índice de columna

# Métodos disponibles por campo clave (Sección 3 del proyecto)
METODOS = {
    1: ["Radix Sort", "Counting Sort"],
    2: ["Bucket Sort"],
    3: ["Quick Sort", "Merge Sort", "Heap Sort"],
}


# ----------------------------------------------------------------- Carga
def cargar_excel(ruta):
    """Carga el Excel a una lista de listas en RAM. Devuelve (datos, ns)."""
    t0 = time.perf_counter_ns()
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    ws = wb.active
    filas = ws.iter_rows(min_row=2, values_only=True)
    datos = [list(f) for f in filas if f and f[0] is not None]
    wb.close()
    return datos, time.perf_counter_ns() - t0


# ------------------------------------------------------------ Quick Sort
def quick_sort(a, k):
    """Quick Sort iterativo, partición en 3 vías (maneja bien llaves repetidas)."""
    a = a[:]
    comp = mov = 0
    pila = [(0, len(a) - 1)]
    while pila:
        lo, hi = pila.pop()
        if lo >= hi:
            continue
        # mediana de tres como pivote
        mid = (lo + hi) // 2
        x, y, z = a[lo][k], a[mid][k], a[hi][k]
        comp += 3
        if (x <= y <= z) or (z <= y <= x):
            pivote = y
        elif (y <= x <= z) or (z <= x <= y):
            pivote = x
        else:
            pivote = z
        lt, i, gt = lo, lo, hi
        while i <= gt:
            comp += 1
            if a[i][k] < pivote:
                a[lt], a[i] = a[i], a[lt]
                mov += 1
                lt += 1
                i += 1
            else:
                comp += 1
                if a[i][k] > pivote:
                    a[i], a[gt] = a[gt], a[i]
                    mov += 1
                    gt -= 1
                else:
                    i += 1
        pila.append((lo, lt - 1))
        pila.append((gt + 1, hi))
    return a, comp, mov


# ------------------------------------------------------------ Merge Sort
def merge_sort(a, k):
    """Merge Sort ascendente (bottom-up, estable)."""
    a = a[:]
    n = len(a)
    comp = mov = 0
    aux = a[:]
    ancho = 1
    while ancho < n:
        for lo in range(0, n, 2 * ancho):
            mid = min(lo + ancho, n)
            hi = min(lo + 2 * ancho, n)
            i, j, p = lo, mid, lo
            while i < mid and j < hi:
                comp += 1
                if a[i][k] <= a[j][k]:
                    aux[p] = a[i]
                    i += 1
                else:
                    aux[p] = a[j]
                    j += 1
                p += 1
                mov += 1
            while i < mid:
                aux[p] = a[i]
                i += 1
                p += 1
                mov += 1
            while j < hi:
                aux[p] = a[j]
                j += 1
                p += 1
                mov += 1
        a, aux = aux, a
        ancho *= 2
    return a, comp, mov


# ------------------------------------------------------------- Heap Sort
def heap_sort(a, k):
    a = a[:]
    n = len(a)
    comp = mov = 0

    def sift_down(inicio, fin):
        nonlocal comp, mov
        raiz = inicio
        while 2 * raiz + 1 <= fin:
            hijo = 2 * raiz + 1
            if hijo + 1 <= fin:
                comp += 1
                if a[hijo][k] < a[hijo + 1][k]:
                    hijo += 1
            comp += 1
            if a[raiz][k] < a[hijo][k]:
                a[raiz], a[hijo] = a[hijo], a[raiz]
                mov += 1
                raiz = hijo
            else:
                return

    for inicio in range((n - 2) // 2, -1, -1):
        sift_down(inicio, n - 1)
    for fin in range(n - 1, 0, -1):
        a[0], a[fin] = a[fin], a[0]
        mov += 1
        sift_down(0, fin - 1)
    return a, comp, mov


# ------------------------------------------------------------ Radix Sort
def radix_sort(a, k):
    """Radix Sort LSD base 10 (cada pasada es un Counting Sort estable)."""
    a = a[:]
    comp = mov = 0
    mayor = max(r[k] for r in a)  # no es una comparación entre registros
    exp = 1
    while mayor // exp > 0:
        conteo = [0] * 10
        for r in a:
            conteo[(r[k] // exp) % 10] += 1
        for d in range(1, 10):
            conteo[d] += conteo[d - 1]
        salida = [None] * len(a)
        for r in reversed(a):
            d = (r[k] // exp) % 10
            conteo[d] -= 1
            salida[conteo[d]] = r
            mov += 1
        a = salida
        exp *= 10
    return a, comp, mov


# --------------------------------------------------------- Counting Sort
def counting_sort(a, k):
    """Counting Sort estable.

    Si el rango (max-min) es razonable usa un arreglo de conteos clásico.
    Los id_embarcacion abarcan ~31,000 millones de valores, así que un arreglo
    no cabe en RAM; en ese caso se usa un diccionario de conteos y se recorren
    únicamente las llaves distintas (en orden con Radix Sort sobre las llaves).
    """
    comp = mov = 0
    menor = min(r[k] for r in a)
    mayor = max(r[k] for r in a)
    comp += 2 * (len(a) - 1)
    rango = mayor - menor + 1
    if rango <= 5_000_000:
        conteo = [0] * rango
        for r in a:
            conteo[r[k] - menor] += 1
        for i in range(1, rango):
            conteo[i] += conteo[i - 1]
        salida = [None] * len(a)
        for r in reversed(a):
            conteo[r[k] - menor] -= 1
            salida[conteo[r[k] - menor]] = r
            mov += 1
        return salida, comp, mov
    # Rango enorme: conteo disperso con diccionario
    conteo = {}
    for r in a:
        conteo[r[k]] = conteo.get(r[k], 0) + 1
    llaves = [[v] for v in conteo]  # ordenar solo las llaves distintas
    llaves, c2, m2 = radix_sort(llaves, 0)
    comp += c2
    mov += m2
    inicio, acum = {}, 0
    for (v,) in llaves:
        inicio[v] = acum
        acum += conteo[v]
    salida = [None] * len(a)
    for r in a:  # pasada estable hacia adelante
        salida[inicio[r[k]]] = r
        inicio[r[k]] += 1
        mov += 1
    return salida, comp, mov


# ----------------------------------------------------------- Bucket Sort
def bucket_sort(a, k):
    """Bucket Sort (acepta enteros o flotantes). Cada cubeta usa inserción."""
    n = len(a)
    comp = mov = 0
    menor = min(r[k] for r in a)
    mayor = max(r[k] for r in a)
    comp += 2 * (n - 1)
    nb = n
    ancho = (mayor - menor) / nb or 1
    cubetas = [[] for _ in range(nb)]
    for r in a:
        i = min(int((r[k] - menor) / ancho), nb - 1)
        cubetas[i].append(r)
        mov += 1
    salida = []
    for c in cubetas:
        for i in range(1, len(c)):  # insertion sort (estable)
            x = c[i]
            j = i - 1
            while j >= 0:
                comp += 1
                if c[j][k] > x[k]:
                    c[j + 1] = c[j]
                    mov += 1
                    j -= 1
                else:
                    break
            c[j + 1] = x
            mov += 1
        salida.extend(c)
    return salida, comp, mov


ALGORITMOS = {
    "Quick Sort": quick_sort,
    "Merge Sort": merge_sort,
    "Heap Sort": heap_sort,
    "Radix Sort": radix_sort,
    "Counting Sort": counting_sort,
    "Bucket Sort": bucket_sort,
}


def ordenar(datos, campo_clave, metodo):
    """Ejecuta y cronometra. Devuelve (ordenado, ns, comparaciones, movimientos)."""
    k = CAMPOS_CLAVE[campo_clave]
    t0 = time.perf_counter_ns()
    res, comp, mov = ALGORITMOS[metodo](datos, k)
    return res, time.perf_counter_ns() - t0, comp, mov


# -------------------------------------------------------------- Búsqueda
def busqueda_lineal(datos, col, valor):
    """Recorre todo el arreglo. Devuelve (registros, filas, ns, comparaciones).
    `filas` son posiciones en RAM empezando en 1."""
    t0 = time.perf_counter_ns()
    comp = 0
    encontrados, filas = [], []
    texto = isinstance(valor, str)
    v = valor.strip().lower() if texto else valor
    for i, r in enumerate(datos):
        comp += 1
        x = r[col]
        if (str(x).strip().lower() == v) if texto else (x == v):
            encontrados.append(r)
            filas.append(i + 1)
    return encontrados, filas, time.perf_counter_ns() - t0, comp


def busqueda_binaria(ordenado, k, valor):
    """Búsqueda binaria sobre la llave ordenada (límite inferior). Devuelve
    (registros, filas, ns, comparaciones) con todos los que tengan esa llave."""
    t0 = time.perf_counter_ns()
    comp = 0
    lo, hi = 0, len(ordenado)
    while lo < hi:
        mid = (lo + hi) // 2
        comp += 1
        if ordenado[mid][k] < valor:
            lo = mid + 1
        else:
            hi = mid
    encontrados, filas = [], []
    i = lo
    while i < len(ordenado):
        comp += 1
        if ordenado[i][k] != valor:
            break
        encontrados.append(ordenado[i])
        filas.append(i + 1)
        i += 1
    return encontrados, filas, time.perf_counter_ns() - t0, comp


# -------------------------------------------------------------- Bitácoras
from datetime import datetime  # noqa: E402

BITACORA_ORDENACION = "bitacora_ordenacion.txt"
BITACORA_BUSQUEDA = "bitacora_busqueda.txt"
REGISTROS_ORDENADOS = "registros_ordenados.txt"


def mensaje_ordenacion(n, campo, metodo, ns, comp, mov):
    return (f"Se ordenaron {n:,} registros por el campo [{campo}] usando el método "
            f"[{metodo}] en [{ns}] nanosegundos. Comparaciones: [{comp}], "
            f"Intercambios/Movimientos: [{mov}]")


def mensaje_busqueda(n, campo, metodo, ns, comp):
    return (f"Se realizó la búsqueda en {n:,} registros por el campo [{campo}] usando "
            f"el método de búsqueda [{metodo}] en [{ns}] nanosegundos. "
            f"Comparaciones: [{comp}]")


def _agregar(ruta, texto):
    with open(ruta, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {texto}\n")


def guardar_bitacora_ordenacion(carpeta, mensaje):
    _agregar(Path(carpeta) / BITACORA_ORDENACION, mensaje)


def guardar_bitacora_busqueda(carpeta, mensaje):
    _agregar(Path(carpeta) / BITACORA_BUSQUEDA, mensaje)


def guardar_registros_ordenados(carpeta, ordenado, campo):
    """Guarda los registros ordenados (se usa en la 1ª ejecución)."""
    with open(Path(carpeta) / REGISTROS_ORDENADOS, "w", encoding="utf-8") as f:
        f.write(f"# {len(ordenado):,} registros ordenados por {campo}\n")
        f.write("\t".join(COLUMNAS) + "\n")
        for r in ordenado:
            f.write("\t".join(str(x) for x in r) + "\n")


def ficha_registro(registro, fila):
    """Texto de 'INFORMACIÓN DEL REGISTRO LOCALIZADO'."""
    raya = "-" * 70
    lineas = [raya, "INFORMACIÓN DEL REGISTRO LOCALIZADO:", raya,
              f"* Fila en memoria RAM: {fila}"]
    lineas += [f"* {c}: {v}" for c, v in zip(COLUMNAS, registro)]
    lineas.append(raya)
    return "\n".join(lineas)