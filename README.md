# Agente SGI con Gemini
Agente web estático (HTML + JavaScript) que lee el Informe de Indicadores en Excel, analiza cumplimiento, tendencia y efectividad de acciones, y exporta Excel, Word y PowerPoint.

## Estructura
- `index.html`, `css/`, `js/app.js`: interfaz y lógica.
- `prompts/system_prompt.md`: instrucciones del agente (edítalo para ajustar el comportamiento).
- `knowledge/`: contexto normativo, de innovación y KPI que se envía a Gemini.
- `data/`: carpeta para tus informes locales. No subas datos reales a un repositorio público.

## Publicar en GitHub Pages
1. Crea un repositorio y sube todo el contenido.
2. Settings → Pages → Deploy from a branch → `main` / root.
3. Abre la URL, pega tu API key (créala en aistudio.google.com/apikey) y carga el Excel.

## Seguridad
- La key queda solo en el navegador (localStorage). Nunca la escribas en el código.
- Con la API gratuita, Google puede usar los datos enviados para mejorar sus productos (verifica los términos vigentes). No cargues información confidencial de ONAC sin autorización; es un tema de 27001.
- Verifica el modelo y los límites gratuitos vigentes en ai.google.dev y ajusta `CFG.model` en `js/app.js`.
