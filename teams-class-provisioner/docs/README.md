README.md
# Teams Class Provisioner

Script en PowerShell para crear automáticamente equipos de clase (EDU_Class) en Microsoft Teams, agregar estudiantes desde un Excel y generar canales privados por estudiante.

## Características
- Crea un Class Team (EDU_Class) directamente (no estándar).
- Pide al usuario activar el equipo en Teams antes de continuar.
- Agrega el owner y estudiantes desde un archivo Excel (.xlsx).
- Crea canales privados individuales por cada estudiante.
- Configurable mediante un archivo `config.txt`.

## Requisitos
- PowerShell 7+ (`pwsh`) en Windows o macOS.
- Módulos:
  - `MicrosoftTeams`
  - `ImportExcel`
- Permisos en el tenant para crear Class Teams (debe ser posible hacerlo manualmente).

Instalar módulos (una vez):
```powershell
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
Install-Module MicrosoftTeams -Scope CurrentUser
Install-Module ImportExcel    -Scope CurrentUser

Estructura del repositorio
teams-class-provisioner/
├─ src/                # Script principal
├─ examples/           # Archivos de ejemplo
├─ docs/               # Documentación
└─ .gitignore

Documentación

QUICKSTART
 – guía rápida (2–3 minutos).

MANUAL
 – manual detallado paso a paso.

TROUBLESHOOTING
 – errores comunes y soluciones.