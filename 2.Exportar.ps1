param(
    [string]$E = $(Get-Location),   # Ruta de destino; por defecto, la ubicación actual
    [switch]$d,                     # Activar modo de depuración con -d
    [switch]$log,                   # Guardar log en la ubicación entregada.
    [string]$FromDate,              # Fecha "DD/MM/YYYY" desde que se comienza a buscar el correo más antiguo, ignorando lo anterior a este parámetro.
    [string]$s,                     # Especifica el store del cual descargar en caso de ser definido.
    [switch]$list                   # lista los stores disponibles.
)

$DebugMode = $d.IsPresent
$ExportMode = $log.IsPresent
$LogFile = Join-Path (Resolve-Path $E).Path "Log.txt"

# Lista de carpetas a excluir, incluyendo "Problemas de sincronización" y "Error local"
$ExcludedFolders = @(
    'Calendario',
    'Cumpleaños',
    'Tareas', 
    'Contactos',
    'Recipient Cache', 
    'Companies', 
    '{A9E2BC46-B3A0-4243-B315-60D991004455}',
    'Organizational Contacts', 
    'PeopleCentricConversation Buddies',
    '{06967759-274D-40B2-A3EB-D7F9E73727D7}', 
    'Journal', 
    'Notas',
    'ExternalContacts', 
    'Problemas de sincronización', 
    'Error local',
    'PersonMetadata',
    'Diario'
)

# Constante para identificar MailItem
$olMailItemClass = 43

# Tamaño máximo del PST en bytes (1.5 GB)
$MaxPSTSizeBytes = 1.5 * 1024 * 1024 * 1024  # Ajusta el tamaño según tus necesidades

#$IgonoredStores = New-Object System.Collections.ArrayList

# Función para manejar mensajes de log con colores y registro en archivo
function Write_Log {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("INFO","SUCCESS","WARNING","ERROR","DEBUG")]
        [string]$Type,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [string]$ForegroundColor  # Parámetro opcional
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"

    # Asignar color predeterminado si no se proporciona
    if (-not $ForegroundColor) {
        switch ($Type) {
            "INFO"     { $ForegroundColor = "White" }
            "SUCCESS"  { $ForegroundColor = "Green" }
            "WARNING"  { $ForegroundColor = "Yellow" }
            "ERROR"    { $ForegroundColor = "Red" }
            "DEBUG"    { $ForegroundColor = "Cyan" }
            default    { $ForegroundColor = "White" }
        }
    }

    # Escribir en la consola si no es DEBUG o si el modo de depuración está activado
    if ($Type -ne "DEBUG" -or $DebugMode) {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }

    # Escribir en el archivo de log
    if ($ExportMode){
        Add-Content -Path $LogFile -Value $logMessage
    }
}

function Handle_Error {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message
    )
    Write_Log -Type "ERROR" -Message "$Message | StackTrace: $($_.Exception.StackTrace)"
}

# Función para determinar el tipo de store
function Get-StoreType {
    param (
        $Store
    )

    switch ($Store.ExchangeStoreType) {
        0 { return "Mailbox" }
        1 { return "Archivo" }
        2 { return "Public" }
        3 { return "Respaldo" }
        default { return "Desconocido" }
    }
}

# Función para obtener información de los archivos de datos
function Get-DataFileInfo{
    param(
        $Namespace
    )

    $Stores = $Namespace.Stores
    $DataFiles = @()

    foreach($store in $Stores){
        if ($null -eq $store){
            continue
        }

        $DisplayName = $store.DisplayName
        $FilePath = $store.FilePath

        $StoreType = Get-StoreType -Store $store

        if($StoreType -eq "Mailbox" -or $StoreType -eq "Archivo"){
            $DataFiles += [PSCustomObject]@{
                DisplayName = $DisplayName
                StoreType = $StoreType
                Store = $store
                FilePath = $FilePath
            }
            Write_Log -Type DEBUG -Message "Encontrado: '$($DisplayName)' | Tipo: $StoreType." -ForegroundColor Green
        }
        else {
            Write_Log -Type DEBUG -Message "Ignorado: '$DisplayName' | Tipo: $($store.ExchangeStoreType)" -ForegroundColor Yellow
        }
    }

    return $DataFiles
}

# Función para sanitizar nombres de archivos y carpetas
function Sanitize_Name {
    param(
        [string]$Name
    )
    return $Name -replace '[<>:"/\\|?*]', ''
}

# Función recursiva para obtener la ruta completa de una carpeta
function Get-FolderPath {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Folder,            # Objeto de la carpeta de Outlook

        [string]$Separator = '/'    # Separador de directorios (por defecto '/')
    )

    # Validación inicial: Verificar que el objeto Folder no sea null
    if ($null -eq $Folder) {
        throw "El objeto Folder proporcionado es null."
    }

    # Caso base: Si no hay padre, devolver el nombre de la carpeta actual
    if ($null -eq $Folder.Parent) {
        return $Folder.Name
    }

    # Intentar obtener el nombre del padre; manejar posibles excepciones
    try {
        #$ParentName = $Folder.Parent.Name
    }
    catch {
        # Si no se puede acceder al padre, asumir que es el nivel raíz
        return $Folder.Name
    }

    # Llamada recursiva para construir la ruta
    return "$(Get-FolderPath -Folder $Folder.Parent -Separator $Separator)$Separator$($Folder.Name)"
}

# Función auxiliar para encontrar la fecha más temprana en una carpeta y sus subcarpetas
function Get-EarliestMailDate {
    param(
        $Folder,
        [string[]]$ExcludedFolders,
        [datetime]$FromDate
    )

    if ($ExcludedFolders -contains $Folder.Name) {
        Write_Log -Type "DEBUG" -Message "Carpeta excluida: '$(Get-FolderPath -Folder $Folder)'."
        return [DateTime]::MaxValue
    }

    $Earliest = [DateTime]::MaxValue

    try {
        # Definir el filtro con el formato de fecha correcto
        $filter = "[ReceivedTime] >= '" + $FromDate.ToString("MM/dd/yyyy HH:mm") + "' AND [MessageClass] = 'IPM.Note'"
        $Items = $Folder.Items.Restrict($filter)
        Write_Log -Type "DEBUG" -Message "Número de elementos filtrados en '$(Get-FolderPath -Folder $Folder)': $($Items.Count)" -ForegroundColor Cyan

        if ($Items.Count -gt 0) {
            try {
                $Items.Sort("[ReceivedTime]", $false)  # Ordenar ascendente (de más antiguo a más reciente)
                $Item = $Items.Item(1)  # Obtener el correo más antiguo
                if ($null -ne $Item -and $null -ne $Item.ReceivedTime) {
                    if ($Item.ReceivedTime -lt $Earliest) {
                        $Earliest = $Item.ReceivedTime
                        Write_Log -Type "DEBUG" -Message "Fecha más temprana en carpeta '$(Get-FolderPath -Folder $Folder)': $($Item.ReceivedTime)"
                    }
                } else {
                    $folderName = if ($Folder.Name) { $Folder.Name } else { "Carpeta Sin Nombre" }
                    Write_Log -Type "WARNING" -Message "Elemento omitido en '$folderName' por no tener 'ReceivedTime'. | MessageClass: $($Item.MessageClass) | Clase: $($Item.Class) | Asunto: $($Item.Subject)"
                }
            } catch {
                $folderName = if ($Folder.Name) { $Folder.Name } else { "Carpeta Sin Nombre" }
                Write_Log -Type "WARNING" -Message "Error al acceder a 'ReceivedTime' en '$folderName': $($_.Exception.Message)"
            }
        }
    } catch {
        Write_Log -Type "ERROR" -Message "Error al procesar la carpeta '$(Get-FolderPath -Folder $Folder)': $($_.Exception.Message)"
    }

    # Recorrer las subcarpetas
    foreach ($SubFolder in $Folder.Folders) {
        $SubEarliest = Get-EarliestMailDate -Folder $SubFolder -ExcludedFolders $ExcludedFolders -FromDate $FromDate
        if ($SubEarliest -lt $Earliest) {
            $Earliest = $SubEarliest
            Write_Log -Type "DEBUG" -Message "Fecha más temprana actualizada desde subcarpeta '$(Get-FolderPath -Folder $SubFolder)': $Earliest"
        }
    }

    return $Earliest
}

# Función para obtener la fecha más temprana de un store
function Get-FirstMailDate {
    param(
        $Store,
        [string[]]$ExcludedFolders,
        [datetime]$FromDate
    )

    $EarliestDate = [DateTime]::MaxValue

    Write_Log -Type "DEBUG" -Message "Procesando store '$($Store.DisplayName)'."
    try {
        $RootFolder = $Store.GetRootFolder()
        $EarliestInStore = Get-EarliestMailDate -Folder $RootFolder -ExcludedFolders $ExcludedFolders -FromDate $FromDate

        if ($EarliestInStore -lt $EarliestDate) {
            $EarliestDate = $EarliestInStore
            Write_Log -Type "DEBUG" -Message "Nueva fecha más temprana encontrada: $EarliestDate en store: '$($Store.DisplayName)'."
        }
    }
    catch {
        Write_Log -Type "ERROR" -Message "Error al procesar el store '$($Store.DisplayName)': $($_.Exception.Message)"
    }

    # Aplicar el FromDate si está definido
    if ($FromDate -ne $null) {
        if ($EarliestDate -lt $FromDate) {
            $EarliestDate = $FromDate
            Write_Log -Type "INFO" -Message "Ajustando la fecha más temprana al parámetro FromDate: $EarliestDate"
        }
    }

    return $EarliestDate
}

# Función para generar rangos de cuatrimestres
function Get-DateRanges {
    param(
        [datetime]$Start,
        [datetime]$End
    )

    $dateRanges = @()

    while ($Start -lt $End){
        $quarterStart = $Start
        $quarterEnd = $Start.AddMonths(3).AddDays(-1)

        if($quarterEnd -gt $End){
            $quarterEnd = $End
        }

        $dateRanges += @($quarterStart,$quarterEnd)
        Write_Log -Type "DEBUG" -Message "Cuatrimestre generado: $quarterStart - $quarterEnd"

        $Start = $Start.AddMonths(3)
    }

    return $dateRanges
}

# Obtener el tamaño del PST.
function Get-PSTSize {
    param(
        [string]$PSTFilePath
    )

    try {
        $FileInfo = Get-Item -Path $PSTFilePath -ErrorAction Stop
        return $FileInfo.Length
    } catch {
        Write_Log -Type "ERROR" -Message "No se pudo obtener el tamaño del archivo PST '$PSTFilePath': $($_.Exception.Message)"
        return 0
    }
}

# Función para cargar los EntryIDs ya exportados
function Load_CopiedEntryIDs {
    param(
        [string]$StoreDisplayName,
        [datetime]$QuarterStartDate,
        [datetime]$QuarterEndDate,
        [string]$E
    )
    
    #$SanitizedStoreName = Sanitize_Name -Name $StoreDisplayName
    $EntryIDsFile = Join-Path $E "$SanitizedStoreName_$($QuarterStartDate.ToString('yyyyMMdd'))_$($QuarterEndDate.ToString('yyyyMMdd')).txt"
    
    if (Test-Path $EntryIDsFile) {
        try {
            $EntryIDs = Get-Content -Path $EntryIDsFile
            return @($EntryIDs)
        } catch {
            Write_Log -Type "ERROR" -Message "No se pudo leer el archivo de EntryIDs '$EntryIDsFile': $($_.Exception.Message)"
            return @()
        }
    } else {
        return @()
    }
}

# Función para guardar los EntryIDs exportados
function Save-CopiedEntryIDs {
    param(
        [string]$StoreDisplayName,
        [datetime]$QuarterStartDate,
        [datetime]$QuarterEndDate,
        [array]$NewEntryIDs,
        [string]$E
    )
    
    #$SanitizedStoreName = Sanitize_Name -Name $StoreDisplayName
    $EntryIDsFile = Join-Path $E "$SanitizedStoreName_$($QuarterStartDate.ToString('yyyyMMdd'))_$($QuarterEndDate.ToString('yyyyMMdd')).txt"
    
    try {
        Add-Content -Path $EntryIDsFile -Value $NewEntryIDs
    } catch {
        Write_Log -Type "ERROR" -Message "No se pudo escribir en el archivo de EntryIDs '$EntryIDsFile': $($_.Exception.Message)"
    }
}

# Función para verificar si el elemento ya existe en el PST
function ElementoYaExiste {
    param(
        [string]$EntryID,
        [array]$CopiedEntryIDs
    )
    
    return $CopiedEntryIDs -contains $EntryID
}

# Función recursiva para copiar ítems a PSTs con validación de EntryID
function Copy-Items {
    param(
        [object]$StoreRootFolder,
        [datetime]$QuarterStartDate,
        [datetime]$QuarterEndDate,
        [string[]]$ExcludedFolders,
        [string]$BaseDataFileName,
        [string]$E,
        [string]$StoreDisplayName
    )
    
    $FolderQueue = New-Object System.Collections.Generic.Queue[Object]
    $FolderQueue.Enqueue($StoreRootFolder)
    
    $PSTPartNumber = 1
    $olStoreUnicode = 3  # Valor para Unicode PST
    
    $TotalItemsToCopy = 0
    $TotalCopiedItems = 0
    $NewCopiedEntryIDs = @()
    
    # Cargar EntryIDs ya exportados
    $CopiedEntryIDs = Load_CopiedEntryIDs -StoreDisplayName $StoreDisplayName `
                                           -QuarterStartDate $QuarterStartDate `
                                           -QuarterEndDate $QuarterEndDate `
                                           -E $E
    
    # Contar el total de elementos a copiar (para el progreso)
    function Count_TotalItems {
        param($Folder)
        $ItemCount = 0
    
        if ($ExcludedFolders -contains $Folder.Name) {
            return 0
        }
    
        # Definir el filtro de fecha y tipo de mensaje
        $filter = "([ReceivedTime] >= '" + $QuarterStartDate.ToString("MM/dd/yyyy HH:mm") + "') AND ([ReceivedTime] <= '" + $QuarterEndDate.ToString("MM/dd/yyyy HH:mm") + "') AND ([MessageClass] = 'IPM.Note')"
    
        try {
            $Items = $Folder.Items.Restrict($filter)
            $ItemCount += $Items.Count
        } catch {
            # Ignorar errores y registrar advertencia
            Write_Log -Type "WARNING" -Message "Error al contar elementos en la carpeta '$($Folder.Name)': $($_.Exception.Message)"
        }
    
        foreach ($SubFolder in $Folder.Folders) {
            $ItemCount += Count_TotalItems -Folder $SubFolder
        }
    
        return $ItemCount
    }
    
    Write_Log -Type "INFO" -Message "Contando el total de elementos a copiar en '$StoreDisplayName' para el cuatrimestre $QuarterStartDate - $QuarterEndDate..."
    $TotalItemsToCopy = Count_TotalItems -Folder $StoreRootFolder
    Write_Log -Type "INFO" -Message "Total de elementos a copiar: $TotalItemsToCopy"
    
    if ($TotalItemsToCopy -eq 0) {
        Write_Log -Type "WARNING" -Message "No se encontraron elementos para copiar en el rango de fechas especificado."
        return
    }
    
    # Funciones internas para manejar PST
    function Open-NewPST {
        param($PSTPartNumber)
    
        if ($PSTPartNumber -eq 1) {
            $ExportFileName = Join-Path (Resolve-Path $E).Path "$BaseDataFileName.pst"
        } else {
            $ExportFileName = Join-Path (Resolve-Path $E).Path "$BaseDataFileName Part $PSTPartNumber.pst"
        }
    
        Write_Log -Type "DEBUG" -Message "ExportFileName: $ExportFileName"
    
        try {
            if (Test-Path $ExportFileName) {
                Write_Log -Type "DEBUG" -Message "Agregando PST existente al perfil de Outlook."
                $Namespace.AddStoreEx($ExportFileName, $olStoreUnicode)
            } else {
                Write_Log -Type "DEBUG" -Message "Creando un nuevo archivo PST."
                $Namespace.AddStoreEx($ExportFileName, $olStoreUnicode)
            }
    
            # Esperar a que el store se agregue
            Start-Sleep -Seconds 5
    
            $PSTStore = $Namespace.Stores | Where-Object { $_.FilePath -ieq $ExportFileName }
            if ($null -ne $PSTStore) {
                Write_Log -Type "DEBUG" -Message "Store PST obtenido: '$($PSTStore.DisplayName)'."
                $PSTRootFolder = $PSTStore.GetRootFolder()
                return @{
                    PSTStore = $PSTStore
                    PSTRootFolder = $PSTRootFolder
                    ExportFileName = $ExportFileName
                }
            } else {
                Write_Log -Type "ERROR" -Message "No se pudo obtener el store del PST '$ExportFileName'."
                return $null
            }
        } catch {
            Write_Log -Type "ERROR" -Message "Error al abrir o crear el PST '$ExportFileName': $($_.Exception.Message)"
            return $null
        }
    }
    
    # Abrir el primer PST
    $PSTInfo = Open-NewPST -PSTPartNumber $PSTPartNumber
    if ($null -eq $PSTInfo) {
        Write_Log -Type "ERROR" -Message "No se pudo abrir o crear el primer PST."
        return
    }
    
    $PSTStore = $PSTInfo.PSTStore
    $PSTRootFolder = $PSTInfo.PSTRootFolder
    $ExportFileName = $PSTInfo.ExportFileName
    
    while ($FolderQueue.Count -gt 0) {
        $CurrentFolder = $FolderQueue.Dequeue()
    
        if ($ExcludedFolders -contains $CurrentFolder.Name) {
            Write_Log -Type "DEBUG" -Message "Carpeta excluida: '$($CurrentFolder.Name)'."
            continue
        }
    
        # Crear la misma estructura de carpetas en el PST
        $FolderPath = Get-FolderPath -Folder $CurrentFolder -Separator '/'
        if ([string]::IsNullOrWhiteSpace($FolderPath)) {
            Write_Log -Type "WARNING" -Message "Ruta de carpeta vacía para la carpeta '$($CurrentFolder.Name)'."
            continue
        }
    
        $FolderPathParts = $FolderPath.Split('/')
        $PSTFolder = $PSTRootFolder
    
        foreach ($Part in $FolderPathParts) {
            if ($Part -eq "") { continue }
            $SubFolder = $PSTFolder.Folders | Where-Object { $_.Name -eq $Part }
            if ($null -eq $SubFolder) {
                try {
                    $SubFolder = $PSTFolder.Folders.Add($Part)
                    Write_Log -Type "DEBUG" -Message "Creada carpeta en PST: '$Part'."
                } catch {
                    Write_Log -Type "ERROR" -Message "Error al crear la carpeta '$Part' en PST: $($_.Exception.Message)"
                    continue  # Saltar a la siguiente carpeta
                }
            }
            $PSTFolder = $SubFolder
        }
    
        # Definir el filtro de fecha y tipo de mensaje
        $filter = "([ReceivedTime] >= '" + $QuarterStartDate.ToString("MM/dd/yyyy HH:mm") + "') AND ([ReceivedTime] <= '" + $QuarterEndDate.ToString("MM/dd/yyyy HH:mm") + "') AND ([MessageClass] = 'IPM.Note')"
        
        # Obtener elementos restringidos por fecha y tipo
        try {
            $Items = $CurrentFolder.Items.Restrict($filter)
            $Items.Sort("[ReceivedTime]", $true)  # Ordenar descendente
        } catch {
            Write_Log -Type "ERROR" -Message "Error al obtener elementos de la carpeta '$($CurrentFolder.Name)': $($_.Exception.Message)"
            continue
        }
    
        foreach ($Item in $Items) {
            if($null -eq $Item.Subject){
                continue 
            }

            if ($null -ne $Item -and $null -ne $Item.ReceivedTime) {
                if ($Item.Class -eq $olMailItemClass -and $Item.MessageClass -eq 'IPM.Note') {
                    # Detalles de depuración
                    if ($DebugMode) {
                        Write_Log -Type "DEBUG" -Message "Detalles del elemento: | EntryID: $($Item.EntryID) | Clase: $($Item.Class) | Asunto: $($Item.Subject) | Fecha de Recepción: $($Item.ReceivedTime)"
                    }

                    # Verificar si el elemento ya existe en el PST
                    if (ElementoYaExiste -EntryID $Item.EntryID -CopiedEntryIDs $CopiedEntryIDs) {
                        Write_Log -Type "DEBUG" -Message "Elemento duplicado ya existe en el PST: '$($Item.Subject)'. Omitiendo copia."
                        continue
                    }

                    try {
                        # Verificar el tamaño actual del PST
                        $CurrentPSTSize = Get-PSTSize -PSTFilePath $ExportFileName
                        if ($CurrentPSTSize -ge $MaxPSTSizeBytes) {
                            # Cerrar el PST actual
                            try {
                                Write_Log -Type "DEBUG" -Message "Tipo de Store: $($PSTStore.GetType().FullName)"
                                $Namespace.RemoveStore($PSTStore.GetRootFolder())
                                Write_Log -Type "DEBUG" -Message "Store '$($PSTStore.DisplayName)' removido correctamente."
                            } catch {
                                Write_Log -Type "ERROR" -Message "Error al remover el store '$($PSTStore.DisplayName)': $($_.Exception.Message)"
                            }

                            # Incrementar el número de parte
                            $PSTPartNumber++
                            Write_Log -Type "INFO" -Message "El PST ha alcanzado el tamaño máximo. Creando nueva parte: $PSTPartNumber."

                            # Abrir un nuevo PST
                            $PSTInfo = Open-NewPST -PSTPartNumber $PSTPartNumber
                            if ($null -eq $PSTInfo) {
                                Write_Log -Type "ERROR" -Message "No se pudo abrir o crear el PST parte $PSTPartNumber."
                                return
                            }

                            $PSTStore = $PSTInfo.PSTStore
                            $PSTRootFolder = $PSTInfo.PSTRootFolder
                            $ExportFileName = $PSTInfo.ExportFileName

                            # Recrear la estructura de carpetas en el nuevo PST
                            $PSTFolder = $PSTRootFolder
                            foreach ($Part in $FolderPathParts) {
                                if ($Part -eq "") { continue }
                                $SubFolder = $PSTFolder.Folders | Where-Object { $_.Name -eq $Part }
                                if ($null -eq $SubFolder) {
                                    try {
                                        $SubFolder = $PSTFolder.Folders.Add($Part)
                                        Write_Log -Type "DEBUG" -Message "Creada carpeta en nuevo PST: '$Part'."
                                    } catch {
                                        Write_Log -Type "ERROR" -Message "Error al crear la carpeta '$Part' en PST: $($_.Exception.Message)"
                                        continue  # Saltar a la siguiente carpeta
                                    }
                                }
                                $PSTFolder = $SubFolder
                            }
                        }

                        # Intentar copiar el elemento con reintentos
                        $maxRetries = 3
                        $retryCount = 0
                        $copied = $false

                        while (-not $copied -and $retryCount -lt $maxRetries) {
                            try {
                                $CopiedItem = $Item.Copy()
                                [void]$CopiedItem.Move($PSTFolder)
                                $TotalCopiedItems++
                                Write_Log -Type "DEBUG" -Message "Copiado: '$($Item.Subject)'."

                                # Agregar EntryID al registro de copiados
                                $NewCopiedEntryIDs += $Item.EntryID
                                $copied = $true
                            } catch {
                                $retryCount++
                                Write_Log -Type "WARNING" -Message "Intento $retryCount de $maxRetries fallido al copiar el elemento '$(Get-FolderPath -Folder $CurrentFolder)/$($Item.Subject)': $($_.Exception.Message)"
                                Start-Sleep -Seconds 2  # Espera antes de reintentar
                            }
                        }

                        if (-not $copied) {
                            Write_Log -Type "ERROR" -Message "No se pudo copiar el elemento '$(Get-FolderPath -Folder $CurrentFolder)/$($Item.Subject)' después de $maxRetries intentos. | MessageClass: $($Item.MessageClass) | Clase: $($Item.Class)"
                        }

                        # Actualizar la barra de progreso
                        if ($TotalItemsToCopy -gt 0) {
                            $ProgressPercent = [int](($TotalCopiedItems / $TotalItemsToCopy) * 100)
                            # Limitar el porcentaje al 100%
                            if ($ProgressPercent -gt 100) { $ProgressPercent = 100 }
                            Write-Progress -Activity "Exportando correos" -Status "Procesando '$($CurrentFolder.Name)': $ProgressPercent% completado" -PercentComplete $ProgressPercent
                        }

                    } catch [System.Runtime.InteropServices.COMException] {
                        Write_Log -Type "ERROR" -Message "Error COM al copiar el elemento '$(Get-FolderPath -Folder $CurrentFolder)/$($Item.Subject)': $($_.Exception.Message) | Código HRESULT: $($_.Exception.ErrorCode) | StackTrace: $($_.Exception.StackTrace)"
                    } catch {
                        Write_Log -Type "ERROR" -Message "Error desconocido al copiar el elemento '$(Get-FolderPath -Folder $CurrentFolder)/$($Item.Subject)': $($_.Exception.Message) | StackTrace: $($_.Exception.StackTrace)"
                    }
                } else {
                    Write_Log -Type "DEBUG" -Message "Elemento omitido en '$(Get-FolderPath -Folder $CurrentFolder)' por no ser un MailItem válido. | MessageClass: $($Item.MessageClass) | Clase: $($Item.Class) | Asunto: $($Item.Subject)"
                }
            } else {
                Write_Log -Type "WARNING" -Message "Elemento omitido en '$(Get-FolderPath -Folder $CurrentFolder)' por no tener 'ReceivedTime'. | MessageClass: $($Item.MessageClass) | Clase: $($Item.Class) | Asunto: $($Item.Subject)"
            }
        }

        # Agregar subcarpetas a la cola
        foreach ($SubFolder in $CurrentFolder.Folders) {
            $FolderQueue.Enqueue($SubFolder)
        }

        # Guardar los nuevos EntryIDs exportados
        if ($NewCopiedEntryIDs.Count -gt 0) {
            Save-CopiedEntryIDs -StoreDisplayName $StoreDisplayName `
                                -QuarterStartDate $QuarterStartDate `
                                -QuarterEndDate $QuarterEndDate `
                                -NewEntryIDs $NewCopiedEntryIDs `
                                -E $E
            # Limpiar la lista de nuevos EntryIDs después de guardarlos
            $NewCopiedEntryIDs = @()
        }
    }
}

# solo listar stores.
if($list.IsPresent){
    try {
        $Outlook = New-Object -ComObject Outlook.Application
        $Namespace = $Outlook.GetNamespace("MAPI")
        $Namespace.Logon($null, $null, $false, $false)
    } catch {
        Handle_Error -Message "No se pudo iniciar la aplicación de Outlook: $($_.Exception.Message)"
        exit
    }
    
    $StoresDF = Get-DataFileInfo -Namespace $Namespace

    foreach ($dataFile in $StoresDF) {
        $Store = $dataFile.Store
        Write_Log -Type "INFO" -Message "$($Store.DisplayName)" -ForegroundColor White
    }

    exit
}

Write_Log -Type INFO -Message "INICIANDO EXPORTADOR v3.1"

# Crear carpeta si no existe.
if (!(Test-Path -Path $E)) {
    try {
        New-Item -ItemType Directory -Path $E -Force | Out-Null
        Write_Log -Type "INFO" -Message "Directorio de destino creado: '$E'." -ForegroundColor Green
    } catch {
        Handle_Error -Message "No se pudo crear el directorio de destino '$E': $($_.Exception.Message)"
        exit
    }
}

# Crear una instancia de Outlook
try {
    Write_Log -Type "INFO" -Message "Iniciando la aplicación de Outlook..." -ForegroundColor White
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $Namespace.Logon($null, $null, $false, $false)
    Write_Log -Type "INFO" -Message "Sesión de Outlook iniciada correctamente." -ForegroundColor Green
} catch {
    Handle_Error -Message "No se pudo iniciar la aplicación de Outlook: $($_.Exception.Message)"
    exit
}

# Convertir $FromDate a DateTime si está definido y no está vacío
if ($FromDate -ne $null -and $FromDate.Trim() -ne "") {
    try {
        $FromDate = [DateTime]::ParseExact($FromDate, "dd/MM/yyyy", $null)
        Write_Log -Type "INFO" -Message "Parámetro FromDate establecido a: $FromDate"
    } catch {
        Handle_Error -Message "El formato de FromDate es inválido. Por favor, use 'DD/MM/YYYY'. Error: $($_.Exception.Message)"
        exit
    }
} else {
    # Si FromDate no está definido, establecer una fecha muy antigua para incluir todos los correos
    $FromDate = [DateTime]::ParseExact("01/01/1900", "dd/MM/yyyy", $null)
    Write_Log -Type "INFO" -Message "Parámetro FromDate no proporcionado. Usando fecha predeterminada: $FromDate"
}

# Obtener información de las tiendas de datos
$StoresDF = Get-DataFileInfo -Namespace $Namespace

# Filtrar los stores si el parámetro $s está definido
if ($null -ne $s -and $s.Trim() -ne "") {
    $StoresDF = $StoresDF | Where-Object { $_.DisplayName -like "*$s*" }
    if ($StoresDF.Count -eq 0) {
        Handle_Error -Message "No se encontró ningún store que coincida con '$s'."
        exit
    } else {
        Write_Log -Type "INFO" -Message "Procesando únicamente el store: '$s'"
    }
}

# Procesar cada store relevante y cada rango de cuatrimestre
foreach ($dataFile in $StoresDF) {
    $Store = $dataFile.Store
    Write_Log -Type "INFO" -Message "Procesando store: '$($Store.DisplayName)'." -ForegroundColor White
    try {
        # Obtener la fecha más temprana del store
        $EarliestDate = Get-FirstMailDate -Store $Store -ExcludedFolders $ExcludedFolders -FromDate $FromDate
        Write_Log -Type "INFO" -Message "Fecha más temprana en store '$($Store.DisplayName)': $EarliestDate" -ForegroundColor Green

        if ($EarliestDate -eq [DateTime]::MaxValue) {
            Write_Log -Type "WARNING" -Message "No se encontraron correos válidos en el store '$($Store.DisplayName)'."
            continue
        }

        # Definir el rango de fechas para los cuatrimestres, desde $EarliestDate hasta hoy
        $StartDate = $EarliestDate
        $EndDate = Get-Date

        # Generar los rangos de cuatrimestres basados en $StartDate y $EndDate
        $dateRanges = Get-DateRanges -Start $StartDate -End $EndDate

        # Iterar sobre los rangos de cuatrimestres y llamar a Copy-Items
        for ($i = 0; $i -lt ($dateRanges.Count / 2); $i++) {
            $QuarterStartDate = $dateRanges[2 * $i]
            $QuarterEndDate = $dateRanges[2 * $i + 1]

            Write_Log -Type "INFO" -Message "Exportando cuatrimestre: $QuarterStartDate - $QuarterEndDate" -ForegroundColor White

            # Determinar el cuatrimestre en base al rango de fechas
            $quarterNumber = [math]::Ceiling($QuarterStartDate.Month / 3)
            $quarter = "Q$($QuarterStartDate.Year)-$quarterNumber"

            # Definir el nombre base del PST
            $BaseDataFileName = Sanitize_Name -Name "$($Store.DisplayName) - $quarter"

            # Llamar a la función Copy-Items
            Copy-Items -StoreRootFolder $Store.GetRootFolder() `
                       -QuarterStartDate $QuarterStartDate `
                       -QuarterEndDate $QuarterEndDate `
                       -ExcludedFolders $ExcludedFolders `
                       -BaseDataFileName $BaseDataFileName `
                       -E $E `
                       -StoreDisplayName $Store.DisplayName
        }
    } catch {
        Handle_Error -Message "Error al procesar el store '$($Store.DisplayName)': $($_.Exception.Message)"
    }
}

# Cerrar la sesión de Outlook
try {
    #$Namespace.Logoff()  # Comentado para evitar cerrar antes de tiempo
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Write_Log -Type "INFO" -Message "Sesión de Outlook cerrada correctamente." -ForegroundColor Green
} catch {
    Write_Log -Type "WARNING" -Message "Error al cerrar la sesión de Outlook: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write_Log -Type "SUCCESS" -Message "Proceso completado exitosamente." -ForegroundColor Green
