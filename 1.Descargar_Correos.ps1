# Cargar la librería de Outlook
#$Outlook = New-Object -ComObject Outlook.Application
#$Namespace = $Outlook.GetNamespace("MAPI")
#$Inbox = $Namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# Obtener todos los correos de la bandeja de entrada
#$AllMails = $Inbox.Items
#$AllMails | Sort-Object ReceivedTime


# Cargar la aplicación de Outlook
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Namespace.Logon($null, $null, $false, $false)  # Iniciar sesión en Outlook

# Función para procesar carpetas recursivamente
function Process-Folder {
    param($Folder)

    # Obtener los elementos de la carpeta
    $Items = $Folder.Items
    if ($Items -ne $null) {
        $TotalItems = $Items.Count
        $ProcessedItems = 0

        foreach ($Item in $Items) {
            try {
                # Verificar si el elemento es un correo (Clase 43)
                if ($Item.Class -eq 43) {
                    # Acceder al asunto para forzar la descarga
                    $Subject = $Item.Subject
                    # Opcional: imprimir el asunto
                    #Write-Host "Asunto: $Subject"
                }
            } catch {
                Write-Host "Error al acceder a un elemento en la carpeta '$($Folder.Name)': $($_.Exception.Message)"
            }

            $ProcessedItems++
            $ProgressPercentage = [math]::Round(($ProcessedItems / $TotalItems) * 100)
            Write-Progress -Activity "Procesando carpeta: $($Folder.Name)" -Status "Procesado $ProcessedItems de $TotalItems elementos" -PercentComplete $ProgressPercentage
        }
    }

    # Procesar subcarpetas
    foreach ($SubFolder in $Folder.Folders) {
        Process-Folder -Folder $SubFolder
    }
}

# Procesar todas las carpetas en todos los almacenes (incluyendo el archivo en línea)
$Stores = $Namespace.Stores
foreach ($Store in $Stores) {
    $RootFolder = $Store.GetRootFolder()
    Write-Host "Procesando almacén: $($Store.DisplayName)"
    Process-Folder -Folder $RootFolder
}

Write-Host "Procesamiento completado."
