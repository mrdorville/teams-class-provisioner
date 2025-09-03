# Teams Class Provisioner

Script en PowerShell para crear equipos de clase (EDU_Class) en Microsoft Teams a partir de listas exportadas desde la PVA (basada en Moodle).

## Características
- Crea un Class Team (plantilla EDU_Class).
- Permite activar el equipo en Teams antes de cargar estudiantes.
- Agrega automáticamente al owner y a los alumnos desde un Excel exportado de la PVA.
- Crea un canal privado por cada estudiante para entregas individuales.
- Modular y configurable con un archivo `config.txt`.

## Documentación
La documentación completa está disponible en la carpeta [`docs/`](./docs/):

- [Guía rápida](./docs/QUICKSTART.md)
- [Manual paso a paso](./docs/MANUAL.md)
- [Solución de problemas](./docs/TROUBLESHOOTING.md)

## Ejemplos
En la carpeta [`examples/`](./examples/) encontrarás:
- `config.txt` de ejemplo con placeholders.
- `alumnos-ejemplo.xlsx` para pruebas (datos ficticios).

## Requisitos
- PowerShell 7 o superior (`pwsh`).
- Módulos:
  ```powershell
  Install-Module MicrosoftTeams -Scope CurrentUser
  Install-Module ImportExcel    -Scope CurrentUser
