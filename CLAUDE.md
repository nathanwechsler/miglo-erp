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

## Backend (Apps Script vía clasp)
- El código del backend vive en apps-script/ (clonado con clasp desde el Script ID 1zEpFmnSPqBB_Ox80Sgpye494BP2MTtAofSN_VLd4COJ6PSBktbhwWTvw).
- Para subir cambios de código: `clasp push` desde apps-script/.
- Para publicar en producción: `clasp deploy --deploymentId AKfycbw3TgY9fSmffqIeuhNgRkcQ-kMdqb2NbpLYlm3aC3fOoqKsRn23Mlh-dQ-4Hp8XBahk4Q` — esto actualiza la versión del deployment EXISTENTE.
- NUNCA correr `clasp deploy` sin --deploymentId (eso crea un deployment nuevo con una URL distinta y rompe el frontend en producción). Si esto llega a pasar por error, avisar inmediatamente sin intentar corregirlo solo.
- Confirmar conmigo antes de cualquier cambio que toque la estructura de Google Sheets (columnas, hojas), permisos de la carpeta de Drive, o el manifest appsscript.json (scopes de OAuth).
- Después de cada cambio de backend: clasp push, luego clasp deploy con el deploymentId de arriba, y recién ahí commit + push a git con mensaje corto en español.

## Notas técnicas
- Funciones de Apps Script terminadas en "_" quedan ocultas del dropdown de ejecución — usar eso para helpers de autorización (ej. autorizarDrive).
- URLs de Apps Script deben usar el formato /macros/s/, no /a/macros/primechef.cl/.
- Hacer commit de este archivo al finalizar.
