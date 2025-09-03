# Teams Class Provisioner

Script en PowerShell para crear equipos de clase (EDU_Class) en Microsoft Teams a partir de listas exportadas desde la PVA (basada en Moodle).

## Características

- Crea un Class Team mediante la plantilla EDU_Class.
- Pausa para que actives el equipo en Teams antes de continuar.
- Agrega automáticamente al owner y a los estudiantes desde un Excel descargado de la PVA.
- Genera un canal privado por cada estudiante para entregas individuales.
- Configurable mediante un archivo `config.txt`.

## Documentación completa

- [Guía rápida](https://github.com/mrdorville/teams-class-provisioner/tree/main/teams-class-provisioner/docs/QUICKSTART.md)  
- [Manual paso a paso](https://github.com/mrdorville/teams-class-provisioner/tree/main/teams-class-provisioner/docs/MANUAL.md)  
- [Solución de problemas (Troubleshooting)](https://github.com/mrdorville/teams-class-provisioner/tree/main/teams-class-provisioner/docs/TROUBLESHOOTING.md)

## Ejemplos incluidos

La carpeta `examples/` contiene:

- `config.txt`: plantilla básica con placeholders que debes completar.  
- `alumnos-ejemplo.xlsx`: archivo de prueba con datos ficticios, siguiendo el formato requerido.

## Requisitos mínimos

- PowerShell 7 o superior (`pwsh`).  
- Los siguientes módulos (una sola vez):
  ```powershell
  Install-Module MicrosoftTeams -Scope CurrentUser
  Install-Module ImportExcel    -Scope CurrentUser
