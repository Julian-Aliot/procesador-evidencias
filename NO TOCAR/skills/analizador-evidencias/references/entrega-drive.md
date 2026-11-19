# Entrega privada en Google Drive

La petición de analizar y entregar autoriza subir el XLSX y el ZIP de resultados a la cuenta de Drive conectada. Comprueba herramientas de subida binaria al comenzar. Si falta capacidad, informa temprano y conserva lo generado para resolver la transferencia.

Usa `Analizador de evidencias / <cliente> / <envío>`. Consulta las carpetas por nombre y padres, verifica sus identificadores y reutiliza únicamente destinos confirmados. Si hay carpetas homónimas ambiguas, pide al usuario identificar el destino. Conserva los identificadores en `NO TOCAR/control/drive.json`, excluido de Git. Cada envío usa una nueva carpeta de ejecución, conservando las anteriores.

La carpeta raíz de este programa es independiente de la del procesador de facturas. Mantén permisos privados y comprueba si el destino tiene permisos heredados compartidos antes de subir datos. Si está compartido, utiliza un destino privado de la cuenta que satisfaga la organización solicitada o aclara la ubicación con el usuario.

Sube el archivo XLSX con MIME application/vnd.openxmlformats-officedocument.spreadsheetml.sheet y el ZIP con application/zip. Conserva el Excel como archivo binario. Sube únicamente los dos archivos de entrega, manteniendo fuera originales adicionales, catálogos y auxiliares.

Antes de repetir una subida incierta, consulta la carpeta y compara sus metadatos con la entrega. Verifica identificador, nombre, tipo, tamaño y, si se expone, checksum de cada archivo subido. Lee sus metadatos antes de declarar entrega completa. Devuelve los enlaces reales observados como Excel en Drive y ZIP en Drive. Si son enlaces de apertura, explica que se descargan desde el menú Descargar de Drive. Para los enlaces relativos hay que descargar y extraer juntos Excel y evidencias.

Ante fallo exclusivo de transferencia, comunica «Procesamiento terminado; entrega por Drive pendiente» con la causa concreta y reintenta únicamente esa transferencia. Un visor local se presenta como tal si llega a utilizarse. Los datos privados permanecen fuera de GitHub.
