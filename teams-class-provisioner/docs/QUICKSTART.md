# Guía Rápida

Este resumen es para que los profesores creen sus equipos de clase en pocos minutos.

## 1. Instalar módulos (una sola vez)
Abrir PowerShell 7 (Windows o macOS):

```powershell
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
Install-Module MicrosoftTeams -Scope CurrentUser
Install-Module ImportExcel    -Scope CurrentUser
2. Preparar los archivos
Copiar examples/config.txt y completarlo:

TeamName = CSTI-1900-5337 - Nombre Asignatura

OwnerUPN = usuario@dominio.edu

Exportar la lista de alumnos a Excel.

Columna C = Correo

Columna D = Nombre

Columna E = Apellidos

Usar examples/alumnos-ejemplo.xlsx como referencia.

3. Ejecutar el script
En la carpeta del proyecto:

powershell
Copy code
pwsh ./src/New-TeamsCourseFromConfig.ps1
El script pedirá:

Seleccionar config.txt

Seleccionar el archivo alumnos.xlsx

4. Activar el equipo
El script se detendrá mostrando un aviso.

Abrir Microsoft Teams → entrar al nuevo equipo → pulsar Activar.

Volver a la consola y presionar Enter.

5. Finalización
El script agregará al owner y estudiantes.

Creará un canal privado por cada estudiante.