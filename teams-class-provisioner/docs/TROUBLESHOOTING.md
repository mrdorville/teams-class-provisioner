# Solución de Problemas

Este documento cubre los errores más comunes al usar el script y cómo resolverlos.

---

## 1. Error: encabezados duplicados en Excel
- El archivo descargado de la PVA puede traer encabezados repetidos.
- El script ya ignora la primera fila (`-NoHeader`), pero asegúrate de no borrar ni modificar los encabezados.
- Verifica que la columna **C** tenga correos válidos.

---

## 2. El equipo no aparece como "Clase"
- El script usa `-Template "EDU_Class"`.
- Si el equipo no tiene pestañas de **Tareas** o **Bloc de notas de clase**, revisa:
  - Confirma que puedes crear un Class Team manualmente en la PVA.
  - Si no, solicita a TI que habiliten licencias de educación en tu cuenta.

---

## 3. No se pueden agregar estudiantes tras crear el equipo
- El equipo queda **inactivo** por defecto.
- Abre Teams y pulsa el botón **Activar** en la parte superior del equipo.
- Una vez activado, vuelve a la consola y presiona **Enter**.

---

## 4. Error al crear canales privados
- Puede ser un problema de propagación: Teams tarda unos segundos en reconocer a un nuevo miembro.
- Ajusta en `config.txt`:
ChannelAddDelaySec = 12
ChannelAddRetries = 10

yaml
Copy code
- Reintenta la ejecución.

---

## 5. Mensaje: "Could not add ... to team"
- Verifica que los correos exportados desde la PVA sean correctos y tengan acceso a Teams.
- Si son cuentas nuevas, puede tardar hasta una hora en propagarse la información en el tenant.

---

## 6. El script no arranca
- Debes ejecutar en PowerShell 7 (`pwsh`), no en la versión clásica de Windows.
- Comando correcto:
```powershell
pwsh ./src/New-TeamsCourseFromConfig.ps1
Si usas Windows, asegúrate de que PowerShell 7 esté instalado desde Microsoft Store.

7. Error al seleccionar archivos
El script abre un cuadro de diálogo (Windows o macOS).

Si falla, también puedes escribir la ruta completa del archivo manualmente.