# Classroom Manager LL. RR.

Herramienta interna en Google Apps Script para administrar Google Classroom desde Google Sheets.

## Flujo de trabajo

Este repositorio se despliega a Google Apps Script mediante GitHub Actions y `clasp`.

Flujo previsto:

```text
ChatGPT / GitHub → GitHub Actions → clasp push → Apps Script → Google Sheets
```

Así se evita copiar y pegar código manualmente en el editor de Apps Script.

## Estado actual

Funciones ya desarrolladas en la versión de trabajo:

- `Start / Setup` para crear hojas base.
- Listado de aulas activas o todas las aulas.
- Procesamiento de alumnos: agregar y eliminar.
- Procesamiento de profesores: agregar y eliminar.
- Consulta de aulas por usuario: alumno o profesor.
- Lista de usuarios del dominio desde Admin Directory.
- Desplegables de usuario en hojas operativas.
- Hojas `LOG` y `ERRORES`.

## Estructura esperada

```text
cm/
├── .github/workflows/
│   ├── import-apps-script.yml
│   └── deploy-apps-script.yml
├── .clasp.example.json
├── .gitignore
├── README.md
└── src/
    ├── Code.js / Código.js
    └── appsscript.json
```

## Secretos requeridos en GitHub Actions

En el repositorio, configurar:

`Settings` → `Secrets and variables` → `Actions` → `New repository secret`

Secretos necesarios:

```text
APPS_SCRIPT_ID
CLASPRC_JSON
```

`APPS_SCRIPT_ID` contiene el ID del proyecto de Apps Script.

`CLASPRC_JSON` contiene el contenido completo del archivo local `.clasprc.json` generado por `clasp login`.

Ese archivo no debe subirse al repositorio.

## Importar el proyecto actual desde Apps Script

Después de configurar los secretos:

1. Ir a `Actions`.
2. Abrir `Import Apps Script`.
3. Pulsar `Run workflow`.

Ese workflow trae el código actual desde Apps Script y lo guarda en `src/`.

## Desplegar cambios hacia Apps Script

Después de que el código esté en GitHub y se hayan hecho cambios:

1. Ir a `Actions`.
2. Abrir `Deploy Apps Script`.
3. Pulsar `Run workflow`.

Ese workflow ejecuta `clasp push --force` y actualiza el proyecto de Apps Script.
