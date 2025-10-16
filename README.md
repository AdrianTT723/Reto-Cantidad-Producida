-Generador Interactivo de Reportes de Producción
Este proyecto consiste en un script de Python, diseñado para ejecutarse en Google Colab, que procesa, visualiza y exporta datos sobre la producción de diversos productos extraídos de la vegetación. El script limpia un archivo Excel inicial, presenta los datos a través de una interfaz interactiva y permite al usuario generar reportes personalizados en formato .xlsx que incluyen tanto los datos tabulados como las gráficas correspondientes.

Características Principales
Limpieza Automática de Datos: Procesa un archivo Excel "sucio", eliminando caracteres no válidos y filas/columnas innecesarias.

Interfaz Interactiva: Utiliza ipywidgets para crear una interfaz amigable donde el usuario puede seleccionar múltiples categorías de productos para analizar.

Visualización Dinámica: Genera gráficos de barras en tiempo real que muestran la evolución de la producción anual para las categorías seleccionadas.

Exportación Personalizada: Permite al usuario nombrar su archivo y exportar un reporte en Excel con dos hojas: una con los datos filtrados y otra con todas las gráficas generadas.

-Prerrequisitos
Para que el script funcione correctamente, debes asegurarte de cumplir con lo siguiente:

Entorno: El script está diseñado para Google Colab.

Archivo de Entrada: Debes tener el archivo de datos tabela289 (2).xlsx en tu Google Drive.

Estructura de Carpetas: El script asume la siguiente estructura de carpetas en tu Google Drive. ¡Es crucial que la respetes!

/MyDrive/
└── RETO CIENCIAS DE DATOS/
    └── tabela289 (2).xlsx
Todos los archivos generados (la base limpia y los reportes) se guardarán en esta misma carpeta.

⚙️ Flujo de Trabajo del Algoritmo
El script está dividido en 6 fases lógicas que se ejecutan secuencialmente para lograr el resultado final.

1. Configuración Inicial e Importación
Importación de Librerías: Se cargan todas las librerías necesarias:

pandas para la manipulación de datos.

google.colab para montar y acceder a Google Drive.

ipywidgets y IPython.display para crear y mostrar la interfaz interactiva.

matplotlib.pyplot para la creación de las gráficas.

openpyxl para escribir los datos y, fundamentalmente, para incrustar imágenes (las gráficas) en el archivo Excel.

os para gestionar archivos del sistema (en este caso, para eliminar las imágenes temporales de las gráficas).

Montaje de Google Drive: Se ejecuta drive.mount() para conectar el entorno de Colab con tu sistema de archivos de Google Drive, permitiendo leer y escribir archivos.

2. Limpieza y Preparación de Datos
Carga de Datos: El script lee el archivo tabela289 (2).xlsx y lo carga en un DataFrame de pandas.

Limpieza de Caracteres: Reemplaza valores no numéricos ("-", "...", ".."), que representan datos faltantes o nulos, con el valor 0.

Eliminación de Metadatos: Se elimina la primera columna y la primera fila del archivo original, ya que contienen información descriptiva que no forma parte de la tabla de datos.

Guardado de Archivo Intermedio: Se crea un nuevo archivo llamado Base_Limpia.xlsx. Este archivo contiene la tabla de datos ya pre-procesada y se guarda sin encabezados para facilitar el siguiente paso.

3. Preparación Final para Gráficas
Lectura del Archivo Limpio: Se carga el archivo Base_Limpia.xlsx.

Asignación de Encabezados: La primera fila de este nuevo archivo (que contiene los nombres de las categorías) se asigna como los encabezados del DataFrame.

Limpieza Final: Se elimina esa primera fila, que ahora es redundante.

Conversión de Tipos: Todas las columnas (excepto la primera, que contiene los años) se convierten a formato numérico. Esto es indispensable para poder realizar cálculos y graficar los datos.

4. Creación de la Interfaz Interactiva
Definición de Categorías: Se crea una lista con todos los nombres de las categorías de productos disponibles para que el usuario pueda seleccionarlas.

Creación de Widgets: Se instancian los tres componentes principales de la interfaz:

SelectMultiple: Un menú desplegable que permite seleccionar una o más categorías de la lista.

Text: Un campo de texto para que el usuario ingrese el nombre deseado para el archivo de reporte.

Button: Un botón con la etiqueta "Exportar a Excel" que activará el proceso de guardado.

Output: Un área designada donde se mostrarán los mensajes y las gráficas generadas.

5. Definición de Funciones de Interacción
Función update_plot():

Activador: Se ejecuta cada vez que el usuario cambia su selección en el menú desplegable.

Acción: Limpia el área de salida y genera un gráfico de barras para cada categoría seleccionada. Cada gráfico muestra la producción a lo largo de los años.

Paso Clave: Antes de mostrar el gráfico en pantalla, lo guarda como un archivo de imagen temporal (.png) en Google Drive. Este paso es crucial para poder incrustarlo en el Excel más adelante.

Función export_to_excel():

Activador: Se ejecuta cuando el usuario hace clic en el botón "Exportar a Excel".

Acción:

Verifica que el usuario haya seleccionado al menos una categoría y haya escrito un nombre para el archivo.

Crea un archivo Excel (.xlsx) usando pd.ExcelWriter con el motor openpyxl.

Hoja 'Datos': Guarda los datos numéricos de las columnas seleccionadas en una hoja llamada Datos.

Hoja 'Gráficas': Crea una segunda hoja llamada Gráficas. Itera sobre las categorías seleccionadas, carga las imágenes .png temporales que se crearon previamente y las inserta una debajo de la otra en esta hoja.

Limpieza Final: Una vez que la imagen ha sido incrustada en el Excel, el archivo .png temporal se elimina de Google Drive para no dejar basura.

Muestra un mensaje de éxito con la ruta del archivo guardado.

6. Vínculo y Visualización
Vinculación de Eventos: Se conectan las funciones a los widgets:

La función update_plot se vincula al evento observe del menú desplegable.

La función export_to_excel se vincula al evento on_click del botón.

Renderizado de la Interfaz: Finalmente, se utiliza display() para mostrar todos los widgets organizados verticalmente en la celda de Colab, listos para que el usuario interactúe con ellos.

-Cómo Usar el Script
Verifica los Prerrequisitos: Asegúrate de que el archivo de entrada y la estructura de carpetas en tu Google Drive sean correctos.

Ejecuta la Celda: Corre la celda completa en tu notebook de Google Colab.

Interactúa:

Selecciona una o más categorías del menú desplegable. Las gráficas aparecerán automáticamente.

Escribe el nombre que deseas para tu reporte en el campo "Nombre archivo".

Haz clic en el botón "Exportar a Excel".

Encuentra tu Reporte: El archivo .xlsx con tu reporte personalizado aparecerá en la carpeta /MyDrive/RETO CIENCIAS DE DATOS/.
