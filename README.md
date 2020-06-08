# SPAMER
Automatic email sender with Gmail IPA and Excel.


## MacOS

### Python
**Para correr el siguiente programa es necesario usar Python 3.6.**
En caso de MacOS Python 3.6 ya viene instalado y se ejecuta corriendo el comando en terminal
``` python3 ```

### Instalar dependencias
Instalar las dependencias con sudo utilizando el siguiente comando.

```sudo pip3 install --upgrade beautifulSoup4 openpyxl google-api-python-client google-auth-httplib2 google-auth-oauthlib```


## Windows

### Python
**Para correr el siguiente programa es necesario usar Python 3.6.**
En caso de Windows es necesario instalarlo, para ello ingresar al siguiente link y descargar el archivo correspondiente para tu sistema.
<https://www.python.org/downloads/release/python-367/>

Tambien se puede incluir Python en las variables de usuario para manejarlo mas facil.

### Instalar dependencias
Instalar las dependencias ejecutar el siguiente comando.

```pip install --upgrade beautifulSoup4 openpyxl google-api-python-client google-auth-httplib2 google-auth-oauthlib```

--------

## Ejecutar

Ejecutar el archivo main.py con Python 3.6 posterior a descargar las dependencias.

## Archivos

Mantener siempre todos los archivos a utilizar dentro de la carpeta root del programa, desde las plantillas hasta los adjuntos.

## Plantilla Excel

SPAMER utiliza una plantilla en Excel la cual cada fila representa un correo distinto a enviar, de modo que el programa itera por todas las filas utilizando el campo EMAIL como el correo destino y la columna ATTACHMENT para los adjuntos del mismo.
A la plantilla se le puede añadir columnas, estas columnas deben contar en el encabezado el nombre de una variable, la cual el programa va a buscar dentro de la plantilla HTML y remplazar por el valor de cada fila en su respectivo correo. 

**IMPORTANTE:** Al correr el programa los campos de la plantilla deben estar en VALORES y no en FORMULAS.

## Gmail API Credenciales

Para correr el programa, es necesario vincular el API a tu cuenta de Google como desarrollador. Para ello seguir los siguientes pasos.

1. Ingresar a <https://developers.google.com/gmail/api/quickstart/python>
2. Hacer click en **Enable the Gmail API**.
3. Seleccionar la configuración **Desktop**.
4. Descargar **credentials.json** haciendo click en **Download client configuration**.
5. Copiar **credentials.json** en el directorio root del programa.

## Plantilla HTML

El programa trabaja con una plantilla HTML la cual puede desarrollar por su cuenta, pero en caso no tenga los conocimientos necesarios puede hacerlo en <https://beefree.io>. Cabe mencionar que **BEEFREE** no sube las imagenes utilizadas en la plantilla a la nube, por lo que al momento de exportar la plantilla es necesario remplazar las direcciones del atributo **src** en las etiquetas **img** por una dirección pública. Esto lo puede hacer subiendo las imagenes a **IMGBB** ingresando al siguiente link <https://imgbb.com>



