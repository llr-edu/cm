# Classroom Manager LL. RR.

Herramienta interna en Google Apps Script para administrar Google Classroom desde Google Sheets.

## Estado actual

Proyecto migrándose a flujo con `clasp` para evitar copiar y pegar código manualmente en Apps Script.

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
├── .clasp.example.json
├── .gitignore
├── README.md
└── src/
    ├── Código.js / Code.js
    └── appsscript.json
```

El archivo real `.clasp.json` no debe subirse al repositorio porque contiene el `scriptId` del proyecto de Apps Script.

## Configuración local

Instalar `clasp`:

```bash
npm install -g @google/clasp
```

Iniciar sesión con la cuenta correcta:

```bash
clasp login --user llrr
```

Crear el archivo local `.clasp.json` a partir del ejemplo:

```bash
cp .clasp.example.json .clasp.json
```

Editar `.clasp.json` y colocar el Script ID real.

Luego traer el código actual del proyecto de Apps Script:

```bash
clasp pull --user llrr
```

Después de revisar que se hayan creado los archivos dentro de `src/`, subirlos a GitHub:

```bash
git add .
git commit -m "Import current Apps Script project"
git push
```

Para enviar cambios desde el repo/local hacia Apps Script:

```bash
clasp push --user llrr
```
