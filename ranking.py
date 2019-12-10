import gspread
import gspread as gs
from oauth2client.service_account import ServiceAccountCredentials

print("Iniciando..")

# Conseguimos las credenciales para usar la API de Google Drive
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
print("Autenticando..")
creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
cliente = gs.authorize(creds)

# Abrimos el meta-archivo y el archivo de Resultados
print("Consiguiendo lista de archivos..")
meta = None
try:
    meta = cliente.open("Meta").sheet1
except gspread.exceptions.SpreadsheetNotFound:
    print("ERROR: No se ha encontrado el archivo de metadatos")
    exit(1)

print("Abriendo archivo de resultados")
resultado = None
try:
    resultado = cliente.open("Resultados").sheet1
except gspread.exceptions.SpreadsheetNotFound:
    print("ERROR: No se ha encontrado el archivo de metadatos")
    exit(1)

# Abrimos los archivos de las evaluaciones, los nombres los da meta
archivos_nombres_meta = meta.col_values(1)[1:]
archivos_tipos_meta = meta.col_values(2)[1:]

num_archivos_solicitados = len(archivos_nombres_meta)
actual = 1

print("Recuperando evaluaciones desde el servidor..")
print("Numero de evaluaciones: %s" % num_archivos_solicitados)

archivos = []
for archivo in archivos_nombres_meta:
    print("Abriendo archivo: %s (%d/%d)" % (archivo, actual, num_archivos_solicitados))
    try:
        archivos.append(cliente.open(archivo).sheet1)
    except gspread.exceptions.SpreadsheetNotFound:
        print("ERROR: No se ha encontrado el archivo")
        exit(1)

    actual += 1

if len(archivos) == 0:
    print("ERROR: No hay ningún archivo para analizar!")
    exit(1)

num_archivos = len(archivos)
print("Archivos recuperados: %d de %d" % (num_archivos, num_archivos_solicitados))

archivos_con_tipo = dict(zip(archivos, archivos_tipos_meta))

num_evaluacion = 1
print("Calculando..")

cache_notas = {}
for archivo, tipo in archivos_con_tipo.items():
    print("%s, %s (%d/%d):" % (archivo.id, tipo, num_evaluacion, num_archivos))

    row_count = archivo.row_count
    for row_num in range(2, row_count):
        row = archivo.row_values(row_num)[1:]
        if not row:
            break

        funcionario = None
        comienzo_notas = None

        if str(tipo) == "AUTO":
            funcionario = row[0]
            comienzo_notas = 1

        puntaje_max = 0
        puntaje = 0

        for espacio in row[comienzo_notas:]:
            if espacio is "" or None:
                break

            puntaje += int(espacio.split(" ")[0])
            puntaje_max += 3

        nota = 10 * (puntaje / puntaje_max)

        print("\tFuncionario: %s, Nota: %.2f" % (funcionario, nota))

        if funcionario in cache_notas.values():
            cache_notas[funcionario] += nota
            break

        cache_notas[funcionario] = nota

print("Guardando resultados..")

resultado.update_cell(1, 1, "Funcionario")
resultado.update_cell(1, 2, "Nota")

cursor_row = 2

print("Escribiendo información para %s funcionarios.." % len(cache_notas.keys()))
for key in sorted(cache_notas, key=cache_notas.get):
    resultado.update_cell(cursor_row, 1, key)
    resultado.update_cell(cursor_row, 2, cache_notas[key])

print("Resultados generados con exito")

