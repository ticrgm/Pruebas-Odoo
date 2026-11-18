# Manual técnico: extracción de datos desde Odoo mediante Python y API JSON-2

## 1. Objetivo

Este documento describe cómo conectarse a Odoo y extraer información de
sus modelos mediante la API JSON-2 utilizando Python. Incluye la
configuración del entorno, la construcción de consultas, el
procesamiento de resultados y la exportación a Excel.

El ejemplo práctico utiliza el modelo `res.partner` para extraer
clientes. La misma estructura puede reutilizarse para consultar otros
modelos, cambiando el modelo, los campos y los filtros.

## 2. Requisitos previos

-   Acceso a una instancia de Odoo.
-   Usuario con permisos de lectura sobre el modelo que se desea
    consultar.
-   API Key válida para la autenticación.
-   Python instalado.
-   CMD o una terminal para ejecutar el programa.

### 2.1. Librerías

Instalar las dependencias dentro del entorno virtual:

``` cmd
pip install requests python-dotenv openpyxl
```

## 3. Configuración del proyecto

Una estructura de proyecto sugerida es:

``` text
ProyectoOdoo/
├── .env
├── main.py
└── README.md
```

### 3.1. Variables de entorno

Crear un archivo `.env` en la carpeta del proyecto:

``` env
ODOO_URL=https://tu-instancia.odoo.com
ODOO_API_KEY=tu_api_key
```

Reemplazar los valores de ejemplo por la dirección de la instancia y la
API Key correspondiente.

**Seguridad:** no publicar ni compartir el archivo `.env`. Tampoco
incluir la API Key directamente en el código fuente.

## 4. Cómo funciona la API JSON-2

Las consultas se realizan mediante una solicitud HTTP al endpoint del
modelo. Para el método `search_read`, la estructura de la URL es:

``` text
https://tu-instancia.odoo.com/json/2/NOMBRE_DEL_MODELO/search_read
```

Por ejemplo, para consultar clientes:

``` text
https://tu-instancia.odoo.com/json/2/res.partner/search_read
```

La autenticación se envía en el encabezado HTTP:

``` python
headers = {
    "Authorization": f"bearer {ODOO_API_KEY}",
    "Content-Type": "application/json"
}
```

La API Key opera con los permisos del usuario al que pertenece. Por
ello, el usuario debe tener acceso de lectura al modelo y a los campos
solicitados.

## 5. Identificación de modelos y campos

Antes de preparar una extracción, identificar:

1.  El modelo que contiene la información.
2.  Los nombres técnicos de los campos disponibles en ese modelo.
3.  Los permisos del usuario de API sobre el modelo y los campos.
4.  Los filtros necesarios para seleccionar los registros.

Algunos modelos habituales son:

  Modelo                Información general
  --------------------- -------------------------------------
  `res.partner`         Contactos, empresas y clientes
  `product.product`     Variantes de productos
  `product.template`    Plantillas de productos
  `account.move`        Documentos contables, como facturas
  `account.move.line`   Líneas de documentos contables
  `stock.move`          Movimientos de inventario
  `mrp.production`      Órdenes de fabricación

Los modelos y campos pueden variar según la versión, los módulos
instalados y las personalizaciones de cada instancia. No asumir que un
campo existe: validarlo antes de incorporarlo a `fields`.

## 6. Construcción de la consulta

La consulta se configura mediante un diccionario Python que se envía en
el cuerpo de la solicitud (`payload`).

Ejemplo:

``` python
payload = {
    "domain": [
        ["customer_rank", ">", 0]
    ],
    "fields": [
        "id",
        "name",
        "vat",
        "email",
        "phone"
    ],
    "limit": 0,
    "order": "name asc"
}
```

### 6.1. Parámetros principales

  -----------------------------------------------------------------------
  Parámetro                           Función
  ----------------------------------- -----------------------------------
  `domain`                            Condiciones que deben cumplir los
                                      registros

  `fields`                            Campos que se desean obtener

  `limit`                             Máximo de registros solicitado; `0`
                                      significa que no se establece un
                                      límite explícito en la petición

  `order`                             Orden de los resultados
  -----------------------------------------------------------------------

**Importante sobre `limit`:** `limit: 0` no garantiza que una consulta
grande devuelva todos los registros. La instancia o el servicio pueden
aplicar límites. Para extracciones extensas se debe implementar
paginación mediante `offset` y `limit`, verificando la cantidad devuelta
en cada bloque.

### 6.2. Ejemplos de condiciones (`domain`)

  Objetivo          Condición
  ----------------- ----------------------------------
  Igual a           `["active", "=", True]`
  Diferente de      `["active", "!=", False]`
  Mayor que         `["customer_rank", ">", 0]`
  Menor que         `["credit_limit", "<", 1000000]`
  Campo no vacío    `["email", "!=", False]`
  Desde una fecha   `["date", ">=", "2026-01-01"]`

Las condiciones pueden combinarse en el dominio. La sintaxis y los
operadores admitidos deben ajustarse al dominio de búsqueda de Odoo.

## 7. Ejecución y respuesta

La solicitud se realiza con `requests.post`. Después de recibir la
respuesta, `raise_for_status()` permite detectar errores HTTP y
`response.json()` convierte la respuesta en estructuras Python:

``` python
response = requests.post(
    url,
    headers=headers,
    json=payload,
    timeout=120
)

response.raise_for_status()
datos = response.json()
```

En una respuesta de `search_read`, normalmente se obtiene una lista de
diccionarios, donde cada diccionario representa un registro.

## 8. Procesamiento de campos relacionales

Los campos relacionales `Many2one` suelen devolverse como una lista con
el identificador y el nombre relacionado, por ejemplo:

``` python
[5, "Transferencia 30 días"]
```

Si para el Excel solo se necesita la descripción, se puede conservar el
segundo elemento:

``` python
for registro in datos:
    for campo, valor in registro.items():
        if isinstance(valor, list):
            registro[campo] = valor[1] if len(valor) > 1 else None
```

Esta transformación facilita la lectura del Excel, pero descarta el
identificador relacionado. Si se necesita utilizar ese ID para cruces
posteriores, conservar ambos valores en lugar de reemplazar la lista.

## 9. Alias para encabezados de Excel

Se puede definir un diccionario para mostrar nombres más comprensibles
en los encabezados, sin cambiar los nombres técnicos utilizados para
consultar Odoo:

``` python
alias_campos = {
    "property_payment_term_id": "Terminos de Pago"
}
```

Al construir los encabezados:

``` python
encabezados = [
    alias_campos.get(columna, columna)
    for columna in columnas
]
```

Si un campo no está definido en `alias_campos`, se utiliza su nombre
original.

## 10. Exportación a Excel

La librería `openpyxl` permite crear un libro, agregar una hoja,
escribir encabezados y registros, ajustar el ancho de las columnas y
habilitar filtros.

El ejemplo incluido en este manual genera el archivo
`Clientes_Odoo.xlsx` con una hoja llamada `Clientes`.

## 11. Ejemplo completo: extracción de clientes (`res.partner`)

El siguiente script reúne los pasos anteriores. Extrae registros
identificados como clientes mediante `customer_rank > 0`, procesa los
campos relacionales y exporta el resultado a Excel.

``` python
import os
import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


# ==================================================
# 1. CONFIGURACIÓN DEL ENTORNO
# ==================================================

# Cargar las variables definidas en el archivo .env.
load_dotenv()

ODOO_URL = os.getenv("ODOO_URL")
ODOO_API_KEY = os.getenv("ODOO_API_KEY")

# Validar que las variables de conexión estén configuradas.
if not ODOO_URL or not ODOO_API_KEY:
    raise ValueError(
        "Falta configurar ODOO_URL u ODOO_API_KEY en el archivo .env"
    )

# Endpoint JSON-2 del modelo res.partner.
url = f"{ODOO_URL.rstrip('/')}/json/2/res.partner/search_read"

# Encabezados de autenticación y formato de la solicitud.
headers = {
    "Authorization": f"bearer {ODOO_API_KEY}",
    "Content-Type": "application/json"
}


# ==================================================
# 2. CONFIGURACIÓN DE ALIAS
# ==================================================

# Los alias solo cambian los encabezados del Excel.
# No modifican los nombres técnicos enviados a Odoo.
alias_campos = {
    "property_payment_term_id": "Terminos de Pago"
}


# ==================================================
# 3. CONFIGURACIÓN DE LA CONSULTA
# ==================================================

# Modelo consultado: res.partner.
#
# DOMAIN:
# customer_rank > 0 selecciona registros que Odoo
# identifica como clientes.
#
# FIELDS:
# Lista de campos que se desean extraer.
#
# LIMIT:
# 0 solicita los registros sin establecer un límite
# explícito. Para grandes volúmenes, implementar
# paginación y verificar los límites de la instancia.
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
        "property_payment_term_id"
    ],
    "limit": 0,
    "order": "name asc"
}


# ==================================================
# 4. EJECUCIÓN DE LA CONSULTA
# ==================================================

try:

    # Enviar la solicitud a Odoo.
    response = requests.post(
        url,
        headers=headers,
        json=payload,
        timeout=120
    )

    # Generar una excepción si Odoo devuelve un error HTTP.
    response.raise_for_status()

    # Convertir la respuesta JSON en una lista de registros.
    datos = response.json()

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

        # Los campos Many2one suelen recibirse como:
        # [ID, Nombre].
        #
        # En este ejemplo se conserva el nombre para
        # facilitar la lectura del archivo Excel.
        #
        # Si se necesita el ID para relaciones posteriores,
        # se debe conservar antes de reemplazar este valor.

        for campo, valor in registro.items():

            if isinstance(valor, list):
                registro[campo] = (
                    valor[1] if len(valor) > 1 else None
                )

        resultados.append(registro)


    # ==================================================
    # 6. CREACIÓN Y EXPORTACIÓN DEL EXCEL
    # ==================================================

    # Crear el libro y asignar nombre a la hoja.
    wb = Workbook()
    ws = wb.active
    ws.title = "Clientes"

    if resultados:

        # Obtener los nombres técnicos de las columnas.
        columnas = list(resultados[0].keys())

        # Aplicar alias cuando exista uno definido.
        # Si no existe, mantener el nombre original.
        encabezados = [
            alias_campos.get(columna, columna)
            for columna in columnas
        ]

        # Escribir los encabezados.
        ws.append(encabezados)

        # Escribir los datos respetando el orden de columnas.
        for registro in resultados:
            ws.append([
                registro.get(columna)
                for columna in columnas
            ])

        # Ajustar el ancho de cada columna según su contenido.
        for columna in ws.columns:

            max_length = max(
                len(str(celda.value or ""))
                for celda in columna
            )

            letra = get_column_letter(columna[0].column)

            ws.column_dimensions[letra].width = min(
                max_length + 2, 40
            )

        # Mantener visible la fila de encabezados.
        ws.freeze_panes = "A2"

        # Habilitar filtros en los encabezados.
        ws.auto_filter.ref = ws.dimensions

    else:
        print("No se encontraron clientes.")

    # Guardar el archivo en el directorio de ejecución.
    archivo = "Clientes_Odoo.xlsx"
    wb.save(archivo)

    print("-----------------------------------")
    print("Exportación finalizada")
    print(f"Archivo generado: {os.path.abspath(archivo)}")
    print(f"Total de clientes: {len(resultados)}")


# ==================================================
# 7. MANEJO DE ERRORES
# ==================================================

except requests.exceptions.HTTPError as error:

    print("Error HTTP al consultar Odoo:")
    print(error)

    # Mostrar el detalle devuelto por Odoo.
    if "response" in locals():
        print(response.text)

except requests.exceptions.RequestException as error:

    print("Error de conexión:")
    print(error)

except Exception as error:

    print("Error durante el procesamiento:")
    print(error)
```

### 11.1. Adaptar el ejemplo a otro modelo

Para consultar otro modelo, modificar principalmente estos elementos:

**A. Endpoint:** reemplazar `res.partner` por el nombre técnico del
modelo requerido.

``` python
url = f"{ODOO_URL.rstrip('/')}/json/2/NOMBRE_MODELO/search_read"
```

**B. Dominio:** establecer los filtros correspondientes al nuevo modelo.

``` python
"domain": [
    ["active", "=", True]
]
```

Este filtro es solo ilustrativo; el campo `active` debe existir en el
modelo consultado.

**C. Campos:** reemplazar la lista `fields` por los campos disponibles
en el modelo seleccionado.

``` python
"fields": [
    "id",
    "name"
]
```

También se debe revisar el procesamiento de los datos y los alias si la
nueva extracción necesita transformaciones diferentes.

## 12. Manejo de errores frecuentes

  Error             Qué revisar
  ----------------- ----------------------------------------------------
  HTTP 401          API Key y autenticación
  HTTP 403          Permisos de acceso del usuario
  HTTP 404          URL, endpoint y modelo
  HTTP 500          Mensaje de error de Odoo y parámetros enviados
  `Invalid field`   Existencia y disponibilidad del campo en el modelo
  Timeout           Tiempo de respuesta y volumen de registros

Si aparece `Invalid field`, revisar que el campo esté escrito
correctamente y que pertenezca al modelo consultado. No basta con que
exista en otro modelo de Odoo.

## 13. Ejecución desde CMD

Ubicarse en la carpeta del proyecto:

``` cmd
cd C:\Proyectos\NombreProyecto
```

Activar el entorno virtual. Si se llama `venv` y está dentro del
proyecto:

``` cmd
venv\Scripts\activate
```

Ejecutar el programa:

``` cmd
python main.py
```

El archivo `Clientes_Odoo.xlsx` se guardará en el directorio de trabajo
desde el que se ejecutó el script.

## 14. Buenas prácticas

-   Consultar solo los campos necesarios para el objetivo de la
    extracción.
-   Validar los modelos y campos en la instancia de Odoo.
-   Utilizar variables de entorno para las credenciales.
-   No compartir archivos `.env` ni API Keys.
-   Considerar paginación para consultas de gran volumen.
-   Registrar el modelo, los filtros y los campos utilizados.
-   Validar la cantidad de registros obtenidos y revisar muestras del
    resultado.
-   Conservar los identificadores cuando se necesiten relaciones o
    cruces entre modelos.
-   Evitar asumir que todos los modelos tienen los mismos campos o
    permisos.

## 15. Resumen del procedimiento

El flujo general de una extracción es:

1.  Configurar la conexión y las credenciales.
2.  Identificar el modelo de Odoo.
3.  Validar los campos disponibles.
4.  Definir el dominio o filtros.
5.  Ejecutar `search_read` mediante JSON-2.
6.  Procesar los valores recibidos.
7.  Exportar los resultados al formato requerido.
8.  Validar el archivo y documentar la extracción.

Este procedimiento sirve como base para desarrollar extracciones de
distintos modelos de Odoo sin tener que construir una conexión nueva
para cada caso.
