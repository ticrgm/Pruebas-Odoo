import os
import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


# ==================================================
# 1. CONFIGURACIÓN DEL ENTORNO
# ==================================================

# Carga las variables de conexión definidas en el archivo .env.
# El archivo debe estar en el directorio desde donde se ejecuta
# este programa, o en una ruta que python-dotenv pueda localizar.
load_dotenv()

ODOO_URL = os.getenv("ODOO_URL")
ODOO_API_KEY = os.getenv("ODOO_API_KEY")

# Validar las variables antes de intentar conectarse a Odoo.
if not ODOO_URL or not ODOO_API_KEY:
    raise ValueError(
        "Falta configurar ODOO_URL u ODOO_API_KEY en el archivo .env"
    )

# Endpoint JSON-2 del modelo res.partner.
# rstrip('/') evita generar una doble barra si ODOO_URL termina en '/'.
url = f"{ODOO_URL.rstrip('/')}/json/2/res.partner/search_read"

# Encabezados utilizados para autenticar la solicitud.
headers = {
    "Authorization": f"bearer {ODOO_API_KEY}",
    "Content-Type": "application/json"
}


# ==================================================
# 2. CONFIGURACIÓN DE ALIAS
# ==================================================

# Permite reemplazar nombres técnicos por encabezados
# más descriptivos en el archivo Excel.
# Los campos no incluidos conservan su nombre original.
alias_campos = {
    "property_payment_term_id": "Terminos de Pago",
    "user_id": "Vendedor",
    "property_product_pricelist": "Tarifa"
}


# ==================================================
# 3. CONFIGURACIÓN DE LA CONSULTA
# ==================================================

# Modelo consultado: res.partner.
#
# DOMAIN:
# customer_rank > 0 filtra los registros que Odoo
# identifica como clientes.
#
# FIELDS:
# Lista de campos que se desean extraer.
# Cada campo debe existir y estar disponible para lectura
# en el modelo res.partner de esta instancia.
#
# LIMIT:
# 0 solicita los registros sin establecer un límite explícito
# en esta petición. La instancia puede aplicar límites propios;
# para grandes volúmenes se debe implementar paginación.
#
# ORDER:
# Ordena los resultados por nombre ascendente.

payload = {
    "domain": [
        ["customer_rank", ">", 0]
    ],
    "fields": [
        "id",
        "name",
        "vat",
        "email",
        "phone",
        "street",
        "street2",
        "city",
        "state_id",
        "zip",
        "country_id",
        "parent_id",
        "customer_rank",
        "active",
        "company_type",
        "fantasy_name",
        "credit_limit",
        "property_payment_term_id",
        "user_id",
        "property_product_pricelist"
    ],
    "limit": 0,
    "order": "name asc"
}


# ==================================================
# 4. EJECUCIÓN DE LA CONSULTA
# ==================================================

try:

    # Enviar la solicitud POST al endpoint de Odoo.
    # El payload contiene el dominio, los campos y el orden.
    response = requests.post(
        url,
        headers=headers,
        json=payload,
        timeout=120
    )

    # Generar una excepción si la respuesta HTTP indica un error.
    response.raise_for_status()

    # Convertir la respuesta JSON en una lista de registros.
    datos = response.json()

    # search_read debe devolver una lista de diccionarios.
    if not isinstance(datos, list):
        raise ValueError(
            "La respuesta de Odoo no tiene el formato de lista esperado."
        )

    print("Conexión exitosa con Odoo")
    print(f"Registros obtenidos: {len(datos)}")


    # ==================================================
    # 5. PROCESAMIENTO DE LOS DATOS
    # ==================================================

    resultados = []

    for registro in datos:

        # Los campos relacionales Many2one suelen recibirse
        # como una lista con esta estructura:
        # [ID, Nombre relacionado]
        #
        # En este ejemplo se conserva solamente el nombre
        # para facilitar la lectura en Excel.
        #
        # Si se necesita el ID para realizar cruces posteriores,
        # se debe conservar por separado antes de reemplazarlo.

        for campo, valor in registro.items():

            if isinstance(valor, list):
                registro[campo] = (
                    valor[1] if len(valor) > 1 else None
                )

        # Incorporar el registro procesado a los resultados.
        resultados.append(registro)


    # ==================================================
    # 6. CREACIÓN Y EXPORTACIÓN DEL EXCEL
    # ==================================================

    # Crear un nuevo libro de Excel y utilizar su hoja activa.
    wb = Workbook()
    ws = wb.active
    ws.title = "Clientes"

    if resultados:

        # Obtener los nombres técnicos de los campos.
        # Se utiliza el primer registro para determinar el orden.
        columnas = list(resultados[0].keys())

        # Aplicar el alias cuando exista.
        # Si no hay alias definido, mantener el nombre técnico.
        encabezados = [
            alias_campos.get(columna, columna)
            for columna in columnas
        ]

        # Escribir los encabezados en la primera fila.
        ws.append(encabezados)

        # Escribir los valores de cada registro respetando
        # el mismo orden de columnas definido anteriormente.
        for registro in resultados:
            ws.append([
                registro.get(columna)
                for columna in columnas
            ])

        # Ajustar el ancho de las columnas según su contenido.
        # Se establece un máximo de 40 caracteres por columna.
        for columna in ws.columns:

            max_length = max(
                len(str(celda.value or ""))
                for celda in columna
            )

            letra = get_column_letter(columna[0].column)

            ws.column_dimensions[letra].width = min(
                max_length + 2, 40
            )

        # Mantener visible la fila de encabezados al desplazarse.
        ws.freeze_panes = "A2"

        # Habilitar filtros en todos los encabezados.
        ws.auto_filter.ref = ws.dimensions

    else:
        print("No se encontraron clientes.")

    # Nombre del archivo Excel generado.
    archivo = "Clientes_Odoo.xlsx"

    # Guardar el archivo en el directorio de trabajo actual.
    wb.save(archivo)

    print("-----------------------------------")
    print("Exportación finalizada")
    print(f"Archivo generado: {os.path.abspath(archivo)}")
    print(f"Total de clientes: {len(resultados)}")


# ==================================================
# 7. MANEJO DE ERRORES
# ==================================================

except requests.exceptions.HTTPError as error:

    # Errores HTTP devueltos por Odoo, por ejemplo:
    # autenticación, permisos, modelo o campo inválido.
    print("Error HTTP al consultar Odoo:")
    print(error)

    # Mostrar el detalle del error entregado por el servidor.
    if "response" in locals():
        print(response.text)


except requests.exceptions.RequestException as error:

    # Errores de conexión, timeout u otros problemas
    # durante la comunicación con el servidor.
    print("Error de conexión:")
    print(error)


except Exception as error:

    # Cualquier otro error durante el procesamiento
    # de los datos o la generación del archivo.
    print("Error durante el procesamiento:")
    print(error)
