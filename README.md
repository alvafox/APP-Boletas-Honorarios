# APP-Boletas-Honorarios

APP para la lectura de boletas de honorarios emitidas por el SII en formato PDF.

La aplicacion puede leer masivamente boletas de honorarios para dejar los datos en una planilla Excel.

En la consola es necesario ejecutar la siguiente linea: pyinstaller --onefile --name "APP BOLETAS (NUEVO)" --hiddenimport win32timezone -F --add-data "Gui.ui;ui" main.py

Luego en la capreta dist estará un archivo autoejecutable que puede buscar y registrar las boletas de honorarios.
