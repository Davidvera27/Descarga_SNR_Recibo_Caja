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





CONSULTAR RECIBO DE PAGO:
Este código de Python está destinado a automatizar el proceso de consulta y descarga de recibos de pago desde un sitio web gubernamental colombiano relacionado con notarías.

1. Cómo funciona:

El código utiliza Selenium, una herramienta para la automatización de navegadores web. Selenium controla un navegador (en este caso, Chrome) para realizar acciones como hacer clic en botones, escribir en campos de texto y extraer información de la página web.
Primero se cargan las librerías necesarias, como Selenium, openpyxl (para trabajar con archivos Excel) y ChromeDriverManager (para manejar el controlador de Chrome).
Luego define funciones para verificar condiciones en filas de un archivo Excel, consultar un número NIR (Número Único de Radicación) en un sitio web, y cerrar el navegador al finalizar.
Posteriormente, se configura Selenium para iniciar una instancia del navegador Chrome maximizado.
Se carga un archivo Excel y se itera sobre las filas, consultando y procesando aquellos registros que cumplan con ciertas condiciones.
Para qué sirve:

El propósito principal es automatizar el proceso de consulta y verificación del estado de recibos de pago asociados a trámites de notarías en Colombia. El código busca automatizar tareas repetitivas que de otra manera serían realizadas manualmente por un usuario.
En qué ámbito es utilizado:

Este tipo de automatización es comúnmente utilizado en notarías donde hay procesos repetitivos que involucran interacción con sitios web. En este caso específico, está relacionado con tareas administrativas o contables dentro de una notaría o incluso dentro de una entidad gubernamental similar en Colombia.

Con qué procesos se relaciona:

Se relaciona con procesos administrativos y contables asociados a notarías o entidades gubernamentales. Específicamente, el código está diseñado para interactuar con el sitio web del Supernotariado de Colombia para consultar el estado de los recibos de pago.

Con qué elementos, archivos y sitios web interactúa:

El código interactúa con archivos Excel (para obtener información de las filas), el sitio web del Supernotariado de Colombia (para consultar el estado de los recibos de pago) y el navegador Chrome (para la automatización de la interacción con el sitio web). Además, parece imprimir resultados en la consola.
En resumen, este script de Python automatiza el proceso de consulta del estado de los recibos de pago asociados a trámites de notarías en Colombia, utilizando Selenium para interactuar con el sitio web del Supernotariado y openpyxl para trabajar con archivos Excel.