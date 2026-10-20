# Informes técnicos Foton — actualización VIN y anexos

## Actualizar la aplicación existente

1. Reemplaza en la raíz del repositorio los archivos streamlit_app.py, base_datos.py, generador.py, plantilla_foton.docx y requirements.txt por los de este paquete.
2. Agrega los archivos anexos.py y acceso.py en la misma carpeta.
3. Conserva los Secrets de Google. No se necesitan Secrets de usuarios ni contraseñas. Mantén streamlit_app.py como archivo principal y reinicia la aplicación.
4. Al conectar, la aplicación agregará las columnas nuevas al final de Informes. No borres ni reordenes las columnas existentes. Si sus nombres no coinciden con la estructura anterior, se detendrá la migración para evitar sobrescribir datos.

No se ha modificado la hoja real desde este paquete. La migración ocurre al ejecutar la aplicación con tus credenciales. Las pestañas utilizadas son Informes y Correlativos; si faltan, se crean. No se altera otra pestaña llamada Datos.

## Nuevos campos

Cliente, km, sucursal y tipo de atención (Postventa o Preventa) aparecen en la identificación del vehículo y se guardan como texto. El kilometraje es un entero sin puntos ni comas; en blanco significa no informado. Estos cuatro campos no se hicieron obligatorios porque no se solicitó.

En EVIDENCIAS es obligatorio cargar una foto legible del VIN (JPG/PNG, hasta 10 MB) y confirmar que corresponde al vehículo. Aparece primero, con descripción explícita. La aplicación no hace reconocimiento automático del VIN ni comprueba visualmente lo declarado por el usuario. Se permiten además hasta 12 fotografías de hallazgos, con descripción obligatoria.

Anexos PDF es opcional: hasta 3 archivos, máximo 20 MB por archivo y 20 páginas en total. Se rechazan archivos ilegibles o con contraseña. Las páginas se convierten localmente en imágenes y se insertan después de la firma; no son texto editable ni conservan firmas digitales, formularios o adjuntos internos del PDF. No se guardan los PDF originales en el Word ni en Sheets. Conserva tus originales si los necesitas.

La base solo guarda textos, nombres de archivos, indicación de foto VIN adjunta y cantidad de páginas PDF. Las columnas agregadas son:
cliente, km, sucursal, tipo_atencion, foto_vin_adjunta, foto_vin_nombre, anexos_pdf, anexos_paginas, usuario_ingreso.
Los registros históricos quedan en blanco en estas columnas; no se les atribuye evidencia inexistente.

## Formato y funcionamiento conservados

Título Informe Técnico de Falla centrado, Calibri 14 y negrita. Textos Calibri 12 justificados. Fotos en recuadros de 14 x 8 cm, sin recorte ni deformación. Firma centrada con nombre en negrita y sin el título RESPONSABLE DEL INFORME. Los anexos siguen a la firma.

Google Sheets funciona como única base. El historial se almacena temporalmente en caché 60 segundos y la preparación de encabezados 5 minutos para reducir lecturas. Las escrituras usan datos frescos e invalidan la caché. Si otro usuario modifica la hoja, espera 60 segundos o pulsa Actualizar historial. La caché reduce consultas, pero no elimina los límites de Google.

Eliminar borra físicamente la fila de Informes. Si se elimina el mayor número existente de ese responsable, el contador retrocede al mayor que queda (o cero si no queda ninguno). Si se elimina uno intermedio, el contador se conserva. No se renumeran informes existentes. No borres la pestaña Correlativos. Los números de informes eliminados podrán reutilizarse, aunque existan Word descargados con esos números. El bloqueo de numeración protege una sola instancia de Streamlit; no ejecutes copias simultáneas del servicio sobre la misma hoja. Ante un fallo de conexión puede quedar un salto en el correlativo.

Las correcciones del historial actualizan solo textos de la base, no un Word descargado. Los archivos solo permanecen temporalmente en la sesión; para generar otro Word debes volver a cargarlos. Descargar el Word es necesario para conservar el informe.

## Verificación realizada

Se comprobó la generación con y sin anexos, foto VIN obligatoria, PDF inválidos y límites, y se revisaron visualmente las páginas Word con PDF vertical y horizontal. La migración y el borrado se verifican con una hoja simulada; la conexión real requiere tus Secrets y permisos en Google Sheets.

Para ejecutar localmente:

    pip install -r requirements.txt
    streamlit run streamlit_app.py

## Integración con tu código actual

streamlit_app.py parte de las 464 líneas que compartiste el 20 de octubre y añade los campos, VIN y PDF. Conserva el texto de Entidades, exportación de historial y respaldo separado, tratamiento de eliminados anteriores, mensajes de errores por operación y limpieza de descargas al editar o borrar. Esta versión incorpora además el ingreso por usuario y el ajuste de correlativos.


## Acceso por usuario sin contraseña

Ingresa solo uno de estos usuarios: avasqueza, ignacio.gutierrez, pgormaz, dimas.toro, desparza o danilo.flores. Se aceptan mayúsculas y espacios al inicio/final. Pulsa Ingresar y el responsable, cargo, correo y prefijo quedan asociados automáticamente. Cambiar usuario limpia los datos y descargas de la sesión anterior.

Esto es identificación por nombre, no autenticación: cualquier persona que conozca un usuario podrá utilizarlo. Todos mantienen acceso al historial compartido, edición, eliminación y respaldo. No se necesitan claves ni el bloque [acceso.usuarios] en Secrets; puedes eliminar ese bloque si lo agregaste, conservando los de Google.

Si ya instalaste el paquete con ingreso por contraseña, reemplaza únicamente acceso.py. streamlit_app.py utiliza la misma función exigir_ingreso y continúa asignando el responsable automáticamente.

## Correlativo después de borrar

Ejemplo: existen DE-0001, DE-0002 y DE-0003. Al eliminar DE-0003, el próximo informe podrá usar DE-0003. Si se elimina DE-0002 y se mantiene DE-0003, el siguiente será DE-0004. Si se eliminan todos los registros de DE, el siguiente será DE-0001. Los registros históricos marcados como eliminados por versiones anteriores siguen ocupando su número mientras su fila exista.

Para corregir un contador adelantado por borrados anteriores: ingresa con el usuario correspondiente, abre Ajustar mi correlativo después de borrados anteriores, marca la confirmación y pulsa Ajustar mi correlativo al historial. Solo recalcula el prefijo del usuario conectado. No cambia los números de registros existentes. Si falla el reajuste después de borrar una fila, la app informa que la fila ya se eliminó e indica usar este ajuste.

Si ya instalaste el paquete de VIN y anexos, reemplaza streamlit_app.py y base_datos.py, agrega acceso.py. No agregues Secrets de usuarios. Los demás archivos no cambian. Para una instalación anterior, utiliza todos los archivos Python, requirements.txt y plantilla_foton.docx del paquete.

Se verificó con pruebas locales: rechazo de usuario desconocido, no lectura del historial antes del ingreso, responsable automático y cambio de usuario y reajuste del contador al borrar último/intermedio/todos. No se modificó ni probó la hoja real de Google Sheets.
