# Convertidor de Actas Viales PRO

Aplicación Windows para asociar documentos Word de consulta (de donde se extraen dominio y número de Acta Vial) con otros documentos Word que se convertirán a PDF.

## Requisitos en la computadora donde se ejecuta
- Windows 10 u 11.
- Microsoft Word de escritorio instalado y activado.
- No requiere Python para usar el ejecutable.

## Uso
1. Abrir la aplicación.
2. Presionar "Agregar Word de consulta" y seleccionar los documentos que contienen dominio y número de acta.
3. Presionar "Agregar Word a convertir" y seleccionar los documentos que se transformarán a PDF.
4. Revisar la asociación sugerida por orden. Doble clic en una fila para cambiar el documento de consulta asociado.
5. Elegir la carpeta de salida.
6. Presionar "Convertir y renombrar PDF".

Los Word originales se abren en modo de solo lectura y no se modifican. Si el nombre PDF ya existe, se agrega un sufijo numérico para no sobrescribirlo.

## Extracción
Busca dominios argentinos actuales (AA123AA) y anteriores (ABC123), y números próximos a etiquetas como "Acta", "Acta Vial" o "N.º de Acta". Revisar siempre la tabla antes de convertir. Si falta alguno de los datos, se bloquea el proceso para evitar nombres incorrectos.

## Ejecutable
Cada push a main/master ejecuta GitHub Actions en Windows y publica el artefacto descargable ConvertidorActasViales-Windows. Entrar en Actions, abrir la última ejecución exitosa y descargar el artefacto.

La compilación no necesita Word instalado en el runner. Microsoft Word sí debe estar instalado en cada PC donde se use el programa.
