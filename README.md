# segy-textfile
Aplicación gráfica para Generar el Encabezado Textual (EBCDIC Header) de Datos Sísmicos 2D Procesados y en formato SEG-Y, según el Manual de Entrega de Información Técnica del Banco de Información Petrolera del Servicio Geológico Colombiano.

Este programa requiere dos archivos auxiliares (ver la carpeta auxFiles):
1.  Archivo Excel con la información de parámetros de adquisición y geofísicos de cada una de las líneas sísmicas, para generar las líneas C1 a C16 del EBCDIC Header del Archivo SEG-Y.
2.  Archivo de texto con la descripción de la posición de grabación de los headers en los diferentes bytes y la secuencia de procesamiento para generar las líneas C17 a C40 del EBCDIC Header del Archivo SEG-Y.

La aplicación le permite al usuario seleccionar el archivo Excel y de texto mencionados anteriormente, la ruta donde van a quedar almacenados los archivos de texto generados, así como el tipo de dato para el cual se van a generar los archivos EBCDIC (PSTM Gathers, Apilado In-In (con filtros y ganancias), Out-Out (sin filtros y sin ganancias), etc).

Para mayor información sobre el estándar exigido por el B.I.P., ver el anexo 1 de  Normatividad para la entrega de información técnica al BIP: https://www2.sgc.gov.co/ProgramasDeInvestigacion/BancoInformacionPetrolera/Paginas/normatividad-entrega-informacion-tecnica-BIP.aspx
