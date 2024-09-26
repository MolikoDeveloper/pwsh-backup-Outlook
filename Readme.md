# Guía para Exportar Correos de Outlook Correctamente
## Requisitos:

[Powershell 7](https://learn.microsoft.com/es-es/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4).

## Primeros Pasos:

Para exportar correos de Outlook correctamente, sigue los pasos a continuación:

- Antes de ejecutar los comandos, abre el archivo `Abrir PowerShell en Bypass Mode.cmd`. Esto permitirá ejecutar los scripts de PowerShell sin restricciones.

# Descargar los Correos:

- Una vez abierto PowerShell en modo Bypass, ejecuta el siguiente comando:

```powershell
.\1.Descargar_Correos.ps1
```
Este comando descargará los correos del perfil de Outlook configurado actualmente. A continuación, se muestra un ejemplo de la salida esperada:

```powershell
PS C:\Users\aarriagadac\Documents\Comandos Outlook> & '.\1-Descargar Correos.ps1'
Procesando almacén: aarriagadac@australis-sa.com
Procesando almacén: Carpetas públicas - aarriagadac@australis-sa.com
Procesando almacén: Archivo de datos de Outlook
Procesamiento completado.
```

## Exportar los Correos:

Después de descargar los correos, ejecuta el siguiente comando para exportarlos:

```powershell
.\2.Exportar.ps1 -E .\aarriagadac
```
#### Resultado esperado:

```powershell
    PS C:\Users\aarriagadac\Documents\Comandos Outlook> .\exportar-debug.ps1 -E .\aarriagadac
    Exportando correos para: Alberto Esteban Arriagada Camblor
    El primer correo fue recibido el: 13/05/2024 10:23:03
    Procesando cuatrimestre: Q2-2024 (13/05/2024 10:23:03 a 12/08/2024 10:23:03)
    Exportando a archivo: C:\Users\aarriagadac\Documents\Comandos Outlook\aarriagadac\Alberto Esteban Arriagada Camblor - Q2-2024.pst
    Copia completada para el cuatrimestre Q2-2024.
    Procesando cuatrimestre: Q3-2024 (13/08/2024 10:23:03 a 14/09/2024 01:58:33)
    Exportando a archivo: C:\Users\aarriagadac\Documents\Comandos Outlook\aarriagadac\Alberto Esteban Arriagada Camblor - Q3-2024.pst
    Copia completada para el cuatrimestre Q3-2024.
    Exportación completada.
```

## Descripción de Argumentos exportar.ps1:

> -E: Define la ubicación en la que deseas exportar los archivos de respaldo. Si no se especifica, los archivos se guardarán en la ubicación actual de la consola.

> -d: Modo de depuración (debug). Activa la visualización detallada de los procesos de exportación, lo cual es útil para el análisis en caso de problemas durante la exportación.

> -log: Guardar log en la carpeta de destino, se generará un log.txt junto a los correos para obtener los detalles de la exportacion.

> -s: especifica el store que quieras descargar.

> -list: lista los stores disponibles en outlook (importante tener la sesion iniciada.)

> -FromDate: filtro desde la fecha en que descargara, formato "DD/MM/YYYY"

## Notas Adicionales:

    Asegúrate de contar con los permisos necesarios para exportar los correos en la ubicación deseada.
    En caso de errores de permisos, verifica los ajustes de seguridad y permisos de escritura en la carpeta de destino.
