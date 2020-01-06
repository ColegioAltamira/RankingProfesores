import statistics
import time

import gspread as gs
from oauth2client.service_account import ServiceAccountCredentials


print("Versión 1.1.0 - camilohernandezcueto@gmail.com")
print("\nIniciando..")

# Conseguimos las credenciales para usar la API de Google Drive
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
print("Autenticando..")
creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
cliente = gs.authorize(creds)

# Abrimos el meta-archivo y el archivo de Resultados
print("Consiguiendo lista de archivos y correos..")
meta = None
emails = None
try:
    meta = cliente.open("Meta").sheet1
    emails = cliente.open("TRABAJADORES 2019-2020").sheet1
except gs.exceptions.SpreadsheetNotFound:
    print("ERROR: No se ha encontrado el archivo de metadatos")
    exit(1)

print("Abriendo archivo de resultados")
resultado = None
try:
    resultado = cliente.open("Resultados").sheet1
except gs.exceptions.SpreadsheetNotFound:
    print("ERROR: No se ha encontrado el archivo de metadatos")
    exit(1)

# Abrimos los archivos de las evaluaciones, los nombres los da meta
archivos_nombres_meta = meta.col_values(1)[1:]
archivos_tipos_meta = meta.col_values(2)[1:]

num_archivos_solicitados = len(archivos_nombres_meta)
actual = 1

print("Recuperando evaluaciones desde el servidor..")

archivos = []
for archivo in archivos_nombres_meta:
    print("Abriendo archivo: %s (%d/%d)" % (archivo, actual, num_archivos_solicitados))
    try:
        archivos.append(cliente.open(archivo).sheet1)
    except gs.exceptions.SpreadsheetNotFound:
        print("ERROR: No se ha encontrado el archivo")
        exit(1)

    actual += 1
    time.sleep(20)

if len(archivos) == 0:
    print("ERROR: No hay ningún archivo para analizar!")
    exit(1)

archivos_con_tipo = dict(zip(archivos, archivos_tipos_meta))

num_archivos = len(archivos)
print("Archivos recuperados: %d de %d" % (num_archivos, num_archivos_solicitados))

num_evaluacion = 1
print("Calculando..")

cache_notas = {}
cache_autoevaluaciones_emails = {}

puntaje = 0
puntaje_max = 0

for archivo, tipo in archivos_con_tipo.items():
    print("%s (%d/%d):" % (archivo.id, num_evaluacion, num_archivos))

    if tipo == "AC" or tipo == "ACM":
        for row_num in range(2, archivo.row_count):
            row = archivo.row_values(row_num)[1:10]
            if not row:
                break

            correo = row[0]
            auto = row[1:]

            puntaje_auto = 0
            puntaje_auto_max = 0
            for nota in auto:
                puntaje_auto += int(nota.split(" ")[0])
                puntaje_auto_max += 4

            nota_auto = puntaje_auto / puntaje_auto_max
            cache_autoevaluaciones_emails[correo] = nota_auto

            print("\tAutoevaluación: %s, Nota: %.2f" % (correo, nota_auto))

        time.sleep(60)

    contador = 1
    nota = 0

    start = 3
    if tipo == "AC":
        start = 11
    if tipo == "ACM":
        start = 15

    nombre_anterior = None
    for col_num in range(start + 1, archivo.col_count):
        col_header = archivo.cell(1, col_num).value
        if col_header == " " or None:
            break

        col_notas = archivo.col_values(col_num)[1:]

        nombre = col_header.split(" [")[0]

        if nombre_anterior is None:
            nombre_anterior = nombre

        if nombre != nombre_anterior:

            if puntaje_max == 0:
                break

            nota = puntaje / puntaje_max
            print("\tFuncionario: %s, Nota: %.2f" % (nombre_anterior, nota))

            time.sleep(30)

            if nombre_anterior in cache_notas:
                cache_notas[nombre_anterior] = statistics.mean([nota, cache_notas[nombre_anterior]])
            else:
                cache_notas[nombre_anterior] = nota

            # Reiniciamos los valores
            nota = 0
            puntaje = 0
            puntaje_max = 0
            nombre_anterior = None

        for espacio in col_notas:

            if espacio == "" or None:
                break

            puntaje += int(espacio.split(" ")[0])
            puntaje_max += 4

    num_evaluacion += 1

print("Interpretando correos:")
cache_autoevaluaciones = {}
for email, nota_emails in cache_autoevaluaciones_emails.items():
    try:
        email_cell = emails.find(email)
    except gs.exceptions.CellNotFound:
        print("\t%s => NOMBRE NO ENCONTRADO" % email)
        continue
    finally:
        time.sleep(30)

    name_cell = emails.cell(email_cell.row, email_cell.col-1)
    name_compress = name_cell.value.lower().replace(" ", "")

    print("\t%s => %s" % (email, name_cell.value))
    cache_autoevaluaciones[name_compress] = nota_emails

print("Uniendo auto y coevaluaciones..")
for key, val in cache_notas.items():
    nota = cache_notas[key]
    name_compress = key.lower().replace(" ", "")

    if name_compress in cache_autoevaluaciones:
        nota = statistics.mean([nota, cache_autoevaluaciones[name_compress]])

    cache_notas[key] = nota

print("Iniciando escritura de resultados..")

resultado.update_cell(1, 1, "Funcionario")
resultado.update_cell(1, 2, "Nota")

cursor_row = 2

print("Escribiendo información de %s funcionarios.." % len(cache_notas.keys()))
for key in sorted(cache_notas, key=cache_notas.get, reverse=True):
    resultado.update_cell(cursor_row, 1, key)
    resultado.update_cell(cursor_row, 2, cache_notas[key])
    cursor_row += 1

print("\nResultados generados con exito")
