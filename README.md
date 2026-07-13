# REINVENTA by Mary Méndez — Sitio web

Sitio web oficial de REINVENTA by Mary Méndez, firma de mentoría y consultoría estratégica de imagen personal.

## Tecnología

Sitio estático de una sola página (HTML + CSS + JavaScript vanilla). Sin framework, sin dependencias, sin pasos de compilación.

## Estructura

```
04_website/
├── index.html                  → Página principal (toda la web)
├── google-sheets-integration.gs → Script para conectar el formulario a Google Sheets
├── robots.txt                  → Instrucciones para motores de búsqueda
├── sitemap.xml                 → Mapa del sitio (actualizar con dominio definitivo)
├── vercel.json                 → Configuración de headers de seguridad y caché
├── .gitignore
├── README.md
└── assets/
    ├── icons/
    │   ├── favicon.svg         → Favicon SVG (fondo plum + isotipo dorado)
    │   ├── favicon-512.png     → Favicon PNG respaldo
    │   ├── apple-touch-icon.png → Ícono para iOS (180×180)
    │   ├── isotipo-dorado.png  → Isotipo original
    │   └── imagotipo-dorado.png → Imagotipo horizontal
    ├── images/                 → Fotografías editoriales (PENDIENTE de entrega)
    └── patterns/               → Patrones de marca (disponibles si se necesitan)
```

## Ejecución local

No requiere instalación. Abre `index.html` directamente en el navegador, o usa cualquier servidor estático:

```bash
# Opción 1: extensión Live Server en VS Code (recomendada)

# Opción 2: Python (si está instalado)
python3 -m http.server 3000
# Abrir http://localhost:3000
```

## Pendientes antes de publicar

| Tarea | Responsable | Estado |
|---|---|---|
| Fotografías editoriales de Mary (hero + sección Mary) | Tu prima / fotógrafo | ⏳ Pendiente |
| Fotografía para sección Talleres | Tu prima / fotógrafo | ⏳ Pendiente |
| URL del Web App de Google Apps Script | Tú | ⏳ Pendiente |
| Usuario de Instagram | Mary / tú | ⏳ Pendiente |
| Correo electrónico de contacto | Mary / tú | ⏳ Pendiente |
| Dominio definitivo | Tú | ⏳ Pendiente |
| Actualizar `canonical`, `og:url`, `robots.txt` Sitemap y `sitemap.xml` con dominio | Tú / Claude | ⏳ Pendiente |
| Imagen Open Graph (1200×630 px, fondo de marca con logo) | Tu prima | ⏳ Pendiente |
| Link del taller (Calendly, WhatsApp o página interna) | Mary / tú | ⏳ Pendiente |
| Aviso de privacidad | Mary / redactor | ⏳ Pendiente |
| Landing page separada (si aplica) | Tu prima | ⏳ No entregada |

## Conectar el formulario a Google Sheets

1. Sigue las instrucciones en `google-sheets-integration.gs`.
2. Al terminar obtendrás una URL que termina en `/exec`.
3. Pégala en `index.html`, línea con `var FORM_ENDPOINT = ''` → `var FORM_ENDPOINT = 'https://script.google.com/macros/s/.../exec'`.

## Subir cambios a GitHub

```bash
git add .
git commit -m "Describe brevemente qué cambiaste"
git push
```

Vercel detecta el push automáticamente y despliega en unos segundos.

## Regresar a una versión anterior

```bash
# Ver historial de commits
git log --oneline

# Regresar temporalmente a una versión anterior (sin borrar historial)
git checkout <id-del-commit>

# Para deshacer el último commit (sin perder los archivos)
git revert HEAD
```

## Archivos que no debes modificar sin revisar antes

- `vercel.json` — cambiar los headers puede afectar seguridad y caché.
- `google-sheets-integration.gs` — cada cambio requiere crear una nueva implementación en Apps Script.
- Las rutas dentro de `assets/icons/` — están referenciadas en el `<head>` de `index.html`.
