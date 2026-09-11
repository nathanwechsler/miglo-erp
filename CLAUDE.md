# Miglo ERP — Reglas del proyecto

## Arquitectura
- Frontend: index.html único (HTML + CSS + JS vanilla, sin frameworks)
- Backend: Google Apps Script como REST API
- Base de datos: Google Sheets (ID: 1hao9nxTbYLNwY7e9ZQy8ElFsJEX-84OkXa2rTGirYOo)
- Almacenamiento de archivos: Google Drive, carpeta "Miglo ERP - Facturas" (1aiQKdQVhUcdGnZnipI78vPkm5b_nW5dx)
- Hosting: GitHub Pages
- Auth: Google OAuth restringido a dominio @primechef.cl (Client ID: 1053887417861-1i2ce27vkhpj633pvrj3gphb6mcugl45)

## Flujo de trabajo
- Después de cada cambio, hacer commit y push automáticamente con un mensaje corto en español describiendo el cambio.
- Antes de ejecutar cambios grandes o que afecten varios módulos, resumir el plan y esperar confirmación.
- Mantener la arquitectura existente: no introducir frameworks ni dependencias nuevas sin avisar.
- Nunca crear un nuevo deployment de Apps Script sin confirmar conmigo primero (cambia la URL y rompe el frontend en producción).

## Manual de marca (Prime Chef)
- Color naranjo: #FE5E19
- Color gris: #DCDCDC
- Color blanco: #FFFFFF
- Tipografía: Inter
- Todo el UI y documentos generados deben respetar estos colores y tipografía.

## Notas técnicas
- Funciones de Apps Script terminadas en "_" quedan ocultas del dropdown de ejecución — usar eso para helpers de autorización (ej. autorizarDrive).
- URLs de Apps Script deben usar el formato /macros/s/, no /a/macros/primechef.cl/.
- Hacer commit de este archivo al finalizar.
