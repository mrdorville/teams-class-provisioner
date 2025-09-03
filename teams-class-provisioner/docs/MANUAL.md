# Manual de Uso Completo

Este documento describe paso a paso cómo usar el script para crear equipos de clase en Microsoft Teams.

---

## 1. Requisitos técnicos
- PowerShell 7 o superior (`pwsh`).
- Conexión a Internet.
- Permisos para crear Class Teams (verificar que manualmente se pueden crear en Teams).

---

## 2. Archivos necesarios

### a) config.txt
Archivo de configuración que indica los datos básicos del equipo.

Ejemplo:
TeamName = CSTI-1900-5337 - Hacking Etico
OwnerUPN = profesor@dominio.edu
Career = ITT

markdown
Copy code

- `TeamName` y `OwnerUPN` son obligatorios.
- `Career` es opcional; si está vacío, el script pedirá el valor.
- `Description` puede dejarse con los placeholders por defecto.

### b) alumnos.xlsx
Archivo con la lista de estudiantes.

#### Cómo obtenerlo desde la PVA
La PVA está basada en Moodle. Para exportar la lista en el formato correcto:

1. Ingresar al curso en la PVA.
2. En la navegación del curso, seleccionar **Calificaciones**.
3. En la barra superior, hacer clic en **Exportar**.
4. Elegir la opción **Hoja de cálculo de Excel**.
5. Descargar el archivo `.xlsx` generado y guardarlo en tu computadora.

Este es el archivo que se usará en el script.

#### Requisitos de contenido
- Columna C = Correo institucional.  
- Columna D = Nombre.  
- Columna E = Apellido(s).  

El resto de columnas pueden estar presentes (el script las ignora).  
Ver `examples/alumnos-ejemplo.xlsx` como referencia.

---

## 3. Ejecución del script
En la carpeta del proyecto:

```powershell
pwsh ./src/New-TeamsCourseFromConfig.ps1
El script pedirá:

Seleccionar el archivo config.txt.

Seleccionar el archivo alumnos.xlsx.

El equipo de clase se creará, pero quedará inactivo.

4. Activar el equipo
Abrir Microsoft Teams.

Entrar al equipo creado.

Pulsar el botón Activar en el banner superior.

Este paso es obligatorio antes de agregar estudiantes o canales.

5. Continuar la ejecución
Volver a la consola de PowerShell.

Presionar Enter para continuar.

El script:

Agregará al owner y a los estudiantes.

Creará un canal privado para cada estudiante.

6. Recomendaciones
No modifiques los nombres de las columnas en el Excel descargado de la PVA.

Si el archivo contiene muchos estudiantes, ajusta los parámetros en config.txt:

ChannelAddDelaySec

ChannelAddRetries

Los nombres de los canales tienen un límite de 50 caracteres. El script los acorta automáticamente si es necesario.