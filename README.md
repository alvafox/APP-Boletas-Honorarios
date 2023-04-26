# APP-Boletas-Honorarios

APP para la lectura de boletas de honorarios emitidas por el SII en formato PDF.

La aplicacion puede leer masivamente boletas de honorarios para dejar los datos en una planilla Excel.

# 0  Clonar el repositorio en la versión de escritorio de Pycharm.

<div align="center">
    <img src="imagenes/Clonar.png" alt="Texto alternativo de la imagen">
</div>

# 1 Luego de clonar el repositorio, es necesario installar las librerias que permitiran usar las funciones del archivo main.py. 

<div align="center">
    <img src="imagenes/Librerias.png" alt="Texto alternativo de la imagen">
</div>

En la consola es necesario ejecutar la siguiente línea: pyinstaller --onefile --name "APP BOLETAS (NUEVO)" --hiddenimport win32timezone -F --add-data "Gui.ui;ui" main.py

Luego en la capreta dist estará un archivo autoejecutable que puede buscar y registrar las boletas de honorarios. Es importante que la aplicación de OUTLOOK este instalada en el escritorio del computador y las boletas deben estar almacenadas en la carpeta por defecto. Por ejemplo, la "Bandeja de entrada". Cualquier otra boleta que no esté dicha carpeta por defecto, no será descargada.

# 2 - Aplicar la línea en la consola.

Una vez clonado el repositorio en Pycharm, se deben instalar todas las librerias asociadas al proyecto.

<div align="center">
    <img src="imagenes/comando.png" alt="Texto alternativo de la imagen">
</div>

# 3 - Se produce la APP en la carpeta "dis".

Seleccionar el rango de fechas para descargar todos los archivos adjuntos en PDF.

<div align="center">
    <img src="imagenes/APP.png" alt="Texto alternativo de la imagen">
</div>

# 4 - Lectura de Boletas

Presionar el botón "Generar EXCEL", mediante esta lectura se genera la planilla para gestionar la información.
<div align="center">
    <img src="imagenes/Generar.png" alt="Texto alternativo de la imagen">
</div>

# 5 - Tabulación de datos

El producto es una planilla Excel, mediante la cual se facilitara la gestion masiva de Boletas de Honorarios.

<div align="center">
    <img src="imagenes/Excel.png" alt="Texto alternativo de la imagen">
</div>
