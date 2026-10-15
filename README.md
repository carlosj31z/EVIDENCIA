# FASC-205 · Evidencia visual de Controles de Cambio

App web estática (Vercel) con Supabase como base de datos y almacenamiento de imágenes.

## Qué hace
- **Importar Excel de CC**: lee el Excel de controles de cambio pendientes (hoja 1, encabezado "CÓDIGO CC"). Une las filas de continuación (celdas combinadas, varias líneas de descripción, varios responsables). Al reimportar, actualiza los datos del Excel sin borrar imágenes ni lo editado; las acciones que ya no figuran se ocultan.
- **Tabla con filtros tipo Excel**: cada columna tiene ▾ con ordenar A→Z / Z→A, buscador, "(Seleccionar todo)" y conteo por valor. Búsqueda global arriba.
- **Panel de la acción**: sección, planta (1/2), anexo y texto de la acción precargados desde el Excel y editables. Pegar imágenes con **Ctrl+V**, arrastrar o elegir archivos. Ordenar, eliminar y **editar** (rotar, recortar, resaltar con recuadro, flecha, deshacer, restaurar original).
- **Vista previa** del formato tal como se imprime (A4 vertical, escala 95 %), página por página.
- **Descargar**: genera `FASC-205-03 EVIDENCIA VISUAL <CÓDIGO> ACCION <N>.xlsx` sobre la plantilla original (logo, encabezado "GARANTÍA DE LA CALIDAD", pie "FASC-205-03", márgenes y área de impresión). Imágenes centradas; 1, 2 o 4 por hoja (automático o elegido). Si hay más hojas, se repite el encabezado del formato (filas 1–17) en cada una.

## Puesta en marcha
1. **Supabase** → SQL Editor → pega y ejecuta `supabase/setup.sql` (crea `cc_acciones`, `cc_evidencias`, el bucket privado `evidencias` y las políticas).
2. `config.js` ya tiene la URL y la anon key del proyecto.
3. **Vercel**: sube esta carpeta a un repositorio de GitHub e impórtalo en Vercel (Framework preset: *Other*, sin build command, output = raíz). También sirve `npx vercel` desde la carpeta.
4. Abre la app → **Importar Excel de CC** → elige el Excel de pendientes.

## Seguridad
Las políticas permiten leer/escribir con la anon key: cualquiera que tenga el link puede usar la app. Para uso interno conviene activar **Vercel → Settings → Deployment Protection** (contraseña o solo miembros del equipo) o agregar login con Supabase Auth y cambiar las políticas a `authenticated`.

## Archivos
- `index.html` · interfaz y estilos
- `app.js` · lógica (importación, filtros, editor, vista previa, generación del .xlsx)
- `assets/template.js` · plantilla FASC-205-03 original y logo (base64)
- `config.js` · conexión a Supabase
- `supabase/setup.sql` · tablas, bucket y políticas
