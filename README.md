# Descarga_SNR_Recibo_Caja
 
Este código de Python es un script que automatiza la tarea de consultar y descargar documentos de un sitio web usando Selenium WebDriver y OpenPyXL para la manipulación de archivos Excel. Aquí está una explicación detallada de cada parte del código:

Propósito del Script:
El script realiza las siguientes acciones:

Inicia sesión en un sitio web.
Lee datos de un archivo Excel.
Para cada fila en el archivo Excel:
Comprueba si ciertas condiciones se cumplen.
Si las condiciones se cumplen, navega a una página web, busca un documento y lo descarga.
Renombra el archivo descargado.
Librerías utilizadas:
openpyxl: Para leer el archivo Excel.
selenium: Para la automatización del navegador web.
os: Para manipulación de archivos y directorios.
time: Para manejar el tiempo de espera.
Detalles del Código:
Configuración del WebDriver y Opciones del Navegador:

Configura el servicio del controlador de Chrome y las opciones del navegador.
Inicio de Sesión:

Abre el navegador y navega a la página de inicio de sesión del sitio web.
Ingresa las credenciales de usuario y contraseña.
Hace clic en el botón "Autenticarse".
Lectura del Archivo Excel:

Lee el archivo Excel que contiene los datos a procesar.
Iteración sobre las Filas del Archivo Excel:

Itera sobre cada fila del archivo Excel.
Verifica si ciertas condiciones se cumplen en la fila.
Si las condiciones se cumplen, procede a buscar y descargar un documento específico.
Búsqueda y Descarga del Documento:

Navega a una página específica del sitio web.
Busca un documento utilizando un valor específico.
Descarga el documento y espera a que se complete la descarga.
Renombra el archivo descargado.
Gestión de Excepciones y Finalización:

Maneja las excepciones que puedan ocurrir durante el proceso.
Cierra el navegador al finalizar el proceso.
Cómo Usar el Script:
Para utilizar este script, se debe modificar ciertas partes, como las rutas de los archivos, las credenciales de inicio de sesión y cualquier otra configuración específica del sitio web o del entorno.

Este script proporciona una forma automatizada de consultar y descargar documentos de un sitio web, lo que puede ser útil para tareas repetitivas o que requieran un gran volumen de interacciones manuales.





