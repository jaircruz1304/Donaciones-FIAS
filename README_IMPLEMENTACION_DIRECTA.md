# FIAS — Control de Donaciones con lectura directa del Excel

## Qué cambia
La versión anterior no consultaba el Excel desde la plataforma. El flujo era:

**Excel SharePoint → GitHub Action cada 30 minutos → JSON versionado → GitHub Pages**

Además, el HTML conservaba un `snapshot` embebido. Eso generaba varias versiones del mismo dato y hacía posible mostrar información anterior cuando fallaba la sincronización.

Esta versión elimina ese esquema. La arquitectura propuesta es:

**Excel SharePoint (fuente única) → API de lectura → Dashboard**

El API descarga el archivo origen cada vez que la plataforma consulta los datos, lo valida, recalcula `Días en gestión` y `Semáforo` y devuelve la información a la web. No se generan archivos JSON permanentes, no se hacen commits automáticos y no existe una copia de datos dentro del HTML.

## Archivos principales
- `Donaciones.html`: dashboard.
- `config.js`: URL del API y frecuencia de actualización.
- `CONTROL_DONACIONES_FIAS_OPTIMIZADO.xlsx`: libro reorganizado para uso como fuente única.
- `api-azure/`: Azure Function que lee el Excel en SharePoint y devuelve JSON en memoria.

## Comportamiento de actualización
- Consulta inicial al abrir la plataforma.
- Botón **Actualizar**: consulta inmediata del Excel origen.
- Actualización automática cada 60 segundos (configurable en `config.js`).
- Si SharePoint o el API fallan, la pantalla NO sustituye los datos con un snapshot antiguo. Si ya existía una consulta exitosa en la sesión, conserva temporalmente esa vista y muestra una alerta visible.
- El botón **Excel contingencia** se mantiene únicamente para una emergencia local; no es el mecanismo normal de actualización.

## Preparación del Excel
El libro optimizado mantiene la hoja `Control` y la tabla `TablaControlDonaciones`. Se corrigieron:
- rangos fijos del resumen;
- total manual que no crecía con la tabla;
- fórmulas de días y semáforo;
- validaciones para estado, respuesta y avance;
- formato condicional para nuevas filas;
- resumen por estados;
- acciones vinculadas a la tabla;
- hoja `Diccionario` para documentar los campos.

El API vuelve a calcular `Días en gestión` y `Semáforo` al leer el libro. Así, la web no depende de los valores almacenados en caché por Excel.

## Implementación recomendada
### 1. Reemplazar/actualizar el Excel origen
Usar `CONTROL_DONACIONES_FIAS_OPTIMIZADO.xlsx` como nuevo archivo fuente. Si se reemplaza el archivo existente conservando el mismo vínculo de SharePoint, no es necesario cambiar la URL de descarga.

### 2. Publicar el API de Azure
Desde la carpeta `api-azure/`, crear una Function App de Node.js y publicar esta carpeta.

Configurar en **Application settings** una de estas dos modalidades:

#### Opción A — enlace de descarga de SharePoint existente
- `SHAREPOINT_DOWNLOAD_URL`: URL directa `download.aspx?share=...`
- `ALLOWED_ORIGINS`: dominio exacto de GitHub Pages, por ejemplo `https://usuario.github.io`

El archivo `local.settings.example.json` contiene la URL que utilizaba el proyecto original como referencia.

#### Opción B — Microsoft Graph (recomendada si el archivo dejará de ser público)
Configurar:
- `TENANT_ID`
- `CLIENT_ID`
- `CLIENT_SECRET`
- `GRAPH_DRIVE_ID`
- `GRAPH_ITEM_ID`
- `ALLOWED_ORIGINS`

Las credenciales permanecen únicamente en el servidor. Nunca deben colocarse en `Donaciones.html` o `config.js`.

### 3. Configurar la web
Editar `config.js`:

```js
window.FIAS_DONACIONES_CONFIG = {
  apiUrl: 'https://NOMBRE-FUNCION.azurewebsites.net/api/donaciones',
  refreshSeconds: 60
};
```

### 4. Publicar en GitHub Pages
Publicar únicamente los archivos estáticos necesarios (`Donaciones.html`, `config.js` y recursos propios si se agregan). Ya no son necesarios:
- `.github/workflows/update-control-donaciones.yml`
- `scripts/build-data.mjs`
- `data/control-donaciones.json`
- `data/control-donaciones-meta.json`

## Regla operativa posterior
La única tarea cotidiana será **actualizar el Excel en SharePoint**. Una vez guardado el cambio, la plataforma lo verá en la próxima consulta automática o al presionar **Actualizar**. No se debe ejecutar ningún proceso de generación o publicación de JSON.
