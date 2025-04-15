Version: 1.0

Autor: Alex Pinedo

Fecha: 14/04/2025

    Descripción: Este script tiene 3 funciones:
    1. Procesar los archivos .txt, extraer datos específicos para ser usados en una plantilla 
        de Word. Luego, convierte el documento a PDF, lo guarda (junto a su txt) en una carpeta de destino.
    2. Realizar una inserción a la tabla "CTABNCO_DOCVIRTUAL" de la BD Oracle, usando los datos 
        extraídos de cada archivo .txt.
    3. (opcional) Si hay archivos .txt que no se pudieron procesar correctamente, se enviara
        un correo, adjuntando los archivos .txt y los errores correspondiente.

    Parametros: 
    Las carpetas y plantilla:
    1. carpeta_origen: Ruta de la carpeta donde se encuentran los archivos .txt a procesar.
    2. carpeta_destino: Ruta de la carpeta donde se guardaran los archivos .pdf y .txt procesados.
    3. carpeta_destino_error: Ruta de la carpeta donde se guardaran los archivos .txt que no  
        se pudieron procesar correctamente.
    4. ruta_plantilla: Ruta de la plantilla de Word que se usara para crear los documentos .docx.

    Conexion a la BD Oracle:
    1. username: Nombre de usuario para la conexión a la BD Oracle.
    2. password: Contraseña para la conexión a la BD Oracle.
    3. dsn: Nombre del servicio de la BD Oracle, colocas la ip y el nombre.

    Configuración del envio de correos:
    1. smtp_server: Servidor SMTP de gmail.
    2. smtp_port: Puerto para SSL.
    3. email_origen: Email que enviara los correos.
    4. email_password: Contraseña de la aplicación del Email.
    5. destinatarios: Lista de Email's de destino.

    Otros:
    1. codificaciones: Lista de codificaciones a probar al abrir los archivos .txt.
    2. tipo: Tipo de archivo a insertar en la tabla "CTABNCO_DOCVIRTUAL". 

    Notas:
    1. La plantilla de Word usa marcadores como (1), (2), etc. para indicar donde se
        deben insertar los datos extraídos del archivo .txt.
        1.1. Para la plantilla se usaron 12 de esos marcadores, de los cuales (3) y (9) son
            obtenidos de la BD Oracle.
    2. Si no hay errores en el proceso, no se enviara ningun correo.
    3. Se configuro "variables de entorno" en mi SO Windows, para la conexion a la BD Oracle 
        y la "clave de aplicacion" del Email.
