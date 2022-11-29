
#region DiskSpd
<#
    Работа с утилитой анализа производительности дисков. Подробнее по ссылке
    https://docs.microsoft.com/en-us/azure-stack/hci/manage/diskspd-overview
#>

<#
.SYNOPSIS
    Командлет загрузки файлов DiskSpd для анализа производительности дисковой подсистемы
    
#>
function Get-DiskSpdFromGitHub {
    
    [CmdletBinding()]
    param (
        [string]$DestinationPath = "$pwd",
        [switch]$DoNotExpandArchive,
        [switch]$DoNotDownloadProcessScript
    )
    
    $client = New-Object System.Net.WebClient -Verbose:$VerbosePreference

    $diskSpdArchivePath = "${DestinationPath}\DiskSpd_latest.zip"
    $diskSpdUrl = "https://github.com/Microsoft/diskspd/releases/latest/download/DiskSpd.zip"

    Write-Verbose "Загрузка файла $diskSpdUrl в $diskSpdArchivePath"
    $client.DownloadFile($diskSpdUrl, $diskSpdArchivePath)

    if ( -not $DoNotExpandArchive ) {
        Expand-Archive -LiteralPath:$diskSpdArchivePath -DestinationPath:$DestinationPath -Verbose:$VerbosePreference
    }

    if ( -not $DoNotDownloadProcessScript ) {

        $processScriptUrl = "https://raw.githubusercontent.com/microsoft/diskspd/master/Process-DiskSpd.ps1"
        $processScriptDestinationPath = "${DestinationPath}\Process-DiskSpd.ps1"

        Write-Verbose "Загрузка файла $processScriptUrl в $processScriptDestinationPath"
        $client.DownloadFile($processScriptUrl, $processScriptDestinationPath)

    }

}

#endregion

<#
.Synopsis
   Очистка временных каталогов 1С
.DESCRIPTION
   Удаляет временные каталоги 1С для пользователя(-ей) с возможностью отбора
.NOTES
   Name: 1CHelper
   Author: yauhen.makei@gmail.com
   https://github.com/emakei/1CHelper.psm1
.EXAMPLE
   # Удаление всех временных каталогов информационных баз для текущего пользователя
   Remove-1CTempDirs
#>
function Remove-1CTempDirs
{
   [CmdletBinding(SupportsShouldProcess = $true)]
   Param
   (
       # Имя пользователя для удаления каталогов(-а)
       [Parameter(Mandatory=$false,
                  ValueFromPipelineByPropertyName=$true,
                  Position=0)]
       [string[]]$User,
       # Фильтр каталогов
       [Parameter(Mandatory=$false,
                  ValueFromPipelineByPropertyName=$true,
                  Position=1)]
       [string[]]$Filter
   )
   
   if( -not $User )
   {
      $AppData = @($env:APPDATA, $env:LOCALAPPDATA)
   }
   else
   {
      Write-Host "Пока не поддерживается"
      return
   }
   
   $Dirs = $AppData | ForEach-Object { Get-ChildItem $_\1C\1cv8*\* -Directory } | Where-Object Name -Match "^\w{8}\-(\w{4}\-){3}\w{12}$"
   if($Filter)
   {
      $Dirs = $Dirs | Where-Object { $_.Name -in $Filter }
   }
   
   if($WhatIfPreference)
   {
      $Dirs | ForEach-Object { "УДАЛЕНИЕ: $($_.FullName)" }
   }
   else
   {
      $Dirs | ForEach-Object { Remove-Item $_.FullName -Confirm:$ConfirmPreference -Verbose:$VerbosePreference -Recurse -Force }
   }
}

<#
.SYNOPSIS
   Преобразует данные файла технологического журнала в таблицу
.DESCRIPTION
   Производит извлечение данных из файла(-ов) технологического журнала и преобразует в таблицу
.NOTES  
    Name: 1CHelper
    Author: yauhen.makei@gmail.com
.LINK  
    https://github.com/emakei/1CHelper.psm1
.INPUTS
   Пусть к файлу(-ам) технологического журнала
.OUTPUTS
   Массив строк технологического журнала
.EXAMPLE
   $table = Get-1CTechJournalLOGtable 'C:\LOG\rmngr_1908\17062010.log'
.EXAMPLE
   $table = Get-1CTechJournalLOGtable 'C:\LOG\' -Verbose
#>
function Get-1CTechJournalLOGtable
{
    [CmdletBinding()]
    [OutputType([Object[]])]
    Param
    (
        # Имя файла лога
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $fileName
    )

    Begin
    {
    $table = @()
    }

    Process
    {

    Get-ChildItem -Path $fileName -Recurse -File | ForEach-Object {
        Write-Verbose $_.FullName
        $creationTime = $_.CreationTime
        $processName = $_.Directory.Name.Split('_')[0]
        $processID = $_.Directory.Name.Split('_')[1]
        Get-1CTechJournalData $_.FullName | ForEach-Object {
            $timeValue = $_.Groups['time'].Value
            $errorTime = $creationTime.AddMinutes($timeValue.Substring(0,2)).AddSeconds($timeValue.Substring(3,2))
            $duration = $timeValue.Split('-')[1]
            $beginTime = $timeValue.Split('.')[1].Split('-')[0]
            $newLine = 1 | Select-Object @{Label='time';           Expression={$errorTime}  }`
                                        ,@{Label='begin';          Expression={$beginTime}  }`
                                        ,@{Label='duration';       Expression={$duration}   }`
                                        ,@{Label='fn:processName'; Expression={$processName}}`
                                        ,@{Label='fn:processID';   Expression={$processID}  }
            $names  = $_.Groups['name'] 
            $values = $_.Groups['value']
            1..$names.Captures.Count | ForEach-Object {
                $propertyName = $names.Captures[$_-1].Value
                $propertyValue = $values.Captures[$_-1].Value
                if ( $null -eq ($newLine | Get-Member $propertyName) ) 
                {
                    Add-Member -MemberType NoteProperty -Name $propertyName -Value $propertyValue -InputObject $newLine 
                }
                else
                {
                    $newValue = @()
                    $newLine.$propertyName | ForEach-Object {$newValue += $_}
                    $newValue += $propertyValue
                    $newLine.$propertyName = $newValue
                }
            }
            $table += $newLine
        }
    }

    }

    End
    {
    $table
    }
}

<#
.Synopsis
   Извлекает данные из xml-файла выгрузки APDEX
.DESCRIPTION
   Производит извлечение данных из xml-файла выгрузки APDEX
.NOTES  
    Name: 1CHelper
    Author: yauhen.makei@gmail.com
.LINK  
    https://github.com/emakei/1CHelper.psm1
.EXAMPLE
   Get-1CAPDEXinfo C:\APDEX\2017-05-16 07-02-54.xml
.EXAMPLE
   Get-1CAPDEXinfo C:\APDEX\ -Verbose
#>
function Get-1CAPDEXinfo
{
    [CmdletBinding()]
    [OutputType([Object[]])]
    Param
    (
        # Имя файла лога
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $fileName
    )

    Begin {
        $xdoc = New-Object System.Xml.XmlDocument
        $tree = @()
    }

    Process {
        Get-ChildItem $fileName -Recurse -File | ForEach-Object {
            Write-Verbose $_.FullName
            try {
                $xdoc.Load($_.FullName)
                if ($xdoc.HasChildNodes) {
                    $tree += $xdoc.Performance.KeyOperation
                }
            } catch {
                $Error | ForEach-Object { Write-Error $_ }
            }
        }
    }

    End {
        $tree
    }
   
}

<#
.SYNOPSIS
   Извлекает данные из файла лога технологического журнала
.DESCRIPTION
   Производит извлечение данных из файла лога технологического журнала
.NOTES  
    Name: 1CHelper
    Author: yauhen.makei@gmail.com
.LINK  
    https://github.com/emakei/1CHelper.psm1
.INPUTS
   Пусть к файлу(-ам) технологического журнала
.OUTPUTS
   Массив данных разбора текстовой информации журнала
.EXAMPLE
   Get-1CTechJournalData C:\LOG\rphost_280\17061412.log
   $properties = $tree | % { $_.Groups['name'].Captures } | select -Unique
.EXAMPLE
   Get-1CTechJournalData C:\LOG\ -Verbose
   $tree | ? { $_.Groups['name'] -like '*Context*' } | % { $_.Groups['value'] } | Select Value -Unique | fl
#>
function Get-1CTechJournalData
{
    [CmdletBinding()]
    [OutputType([Object[]])]
    Param
    (
        # Имя файла лога
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $path
    )

    Begin {
        # Данный шаблон соответствует одной записи журнала
        $template = @"
^(?<line>(?<time>\d\d\:\d\d\.\d{6}\-\d)\,(?<type>\w+)\,(?<level>\d)(\,(?<name>(\w+\:)?\w+)\=(?<value>([^"'\,\n]+|(\"[^"]+\")|(\'[^']+\'))?))+.*)
"@
        $regex = New-Object System.Text.RegularExpressions.Regex ($template, [System.Text.RegularExpressions.RegexOptions]::Multiline)
        $tree = @()
    }

    Process {
        Get-ChildItem $path -Recurse -File | ForEach-Object {
            Write-Verbose $_.FullName
            $rawText = Get-Content $_.FullName -Encoding UTF8 -Raw
            if ($rawText) {
                $textMatches = $regex.Matches($rawText)
                $tree += $textMatches
            }
        }
    }

    End {
        $tree
    }
   
}


<#
.SYNOPSIS
    Возвращает данные счётчиков производительности (рекомендуемые в методическом пособии по эксплуатации крупных информационных систем на платформе 1С:Предприятие 8)
    см. также https://docs.microsoft.com/en-us/troubleshoot/sql/performance/troubleshoot-sql-io-performance#graphical-representation-of-the-methodology
#>
function Get-PerfCounters {
    
    param (

        # Имя компьютера(-ов) для сбора данных счётчиков производительности
        [string[]]$ComputerName,

        # Продолжать непрерывно (пока явно не будет отменено Ctrl+C)
        [switch]$Continuous,

        # Определяет количество замеров для каждого из счётчиков до остановки сбора данных
        [int64]$MaxSamples,

        # Определяет интервал времени между замерами в секундах
        [int32]$SampleInterval = 1,

        # Получить данные всех рекомендуемых счетчиков производительности (из методического пособия по эксплуатации крупных инф. систем)
        [switch]$AllRecomended,

        # Получить данные счетчиков производительности памяти
        [switch]$Memory,

        # Получить данные счетчиков производительности логических дисков
        [switch]$LogicalDisk,

        # Получить данные счетчиков производительности процессора
        [switch]$Processor,

        # Получить данные счетчиков производительности сетевых интерфейсов
        [switch]$NetworkInterface,

        # Получить данные счетчиков производительности физических дисков
        [switch]$PhysicalDisk,

        # Получить данные счетчиков производительности файла подкачки
        [switch]$PagingFile,

        # Получить данные счетчиков производительности SQL сервера
        [switch]$SQLServer

    )

    $counters = @()
    
    if ( $Memory -or $AllRecomended ) {
        
        $counters += '\Memory\Available Mbytes'
        $counters += '\Memory\Pages/sec'

    }

    if ( $LogicalDisk -or $AllRecomended ) {

        $counters += '\LogicalDisk(_Total)\Free Megabytes'
        $counters += '\LogicalDisk(*)\% Disk Time'
        $counters += '\LogicalDisk(*)\% Idle Time' 
        $counters += '\LogicalDisk(*)\% Disk Write Time'
        $counters += '\LogicalDisk(*)\% Disk Read Time'
    }

    if ( $Processor -or $AllRecomended ) {

        $counters += '\Processor(_Total)\% Processor Time'
        $counters += '\System\Processor Queue Length'

    }

    if ( $PhysicalDisk -or $AllRecomended ) {
        
        $counters += '\PhysicalDisk(_Total)\Avg. Disk Queue Length'
        $counters += '\PhysicalDisk(*)\Avg. Disk Queue Length'
        $counters += '\PhysicalDisk(_Total)\Avg. Disk sec/Read'
        $counters += '\PhysicalDisk(_Total)\Avg. Disk sec/Write'
        $counters += '\PhysicalDisk(*)\% Disk Time'
        $counters += '\PhysicalDisk(*)\% Idle Time'
        $counters += '\PhysicalDisk(*)\% Disk Write Time'
        $counters += '\PhysicalDisk(*)\% Disk Read Time'

    }

    if ( $PagingFile ) {

        $counters += '\Paging File(_Total)\*'

    }

    if ( $SQLServer ) {

        $counters += '\SQLServer:Access Methods\Full Scans/sec'
        $counters += '\SQLServer:Buffer Manager\Buffer cache hit ratio'
        $counters += '\SQLServer:Buffer Manager\Free list stalls/sec'
        $counters += '\SQLServer:Buffer Manager\Lazy writes/sec'
        $counters += '\SQLServer:Buffer Manager\Page life expectancy'
        $counters += '\SQLServer:Databases(_Total)\Active Transactions'
        $counters += '\SQLServer:Databases(_Total)\Transactions/sec'
        $counters += '\SQLServer:General Statistics\Active Temp Tables'
        $counters += '\SQLServer:General Statistics\Transactions' 
        $counters += '\SQLServer:Locks(_Total)\Average Wait Time (ms)'
        $counters += '\SQLServer:Locks(_Total)\Lock Requests/sec'
        $counters += '\SQLServer:Locks(_Total)\Lock Timeouts (timeout > 0)/sec'
        $counters += '\SQLServer:Locks(_Total)\Lock Timeouts/sec'
        $counters += '\SQLServer:Locks(_Total)\Lock Wait Time (ms)'
        $counters += '\SQLServer:Locks(_Total)\Lock Waits/sec'
        $counters += '\SQLServer:Locks(_Total)\Number of Deadlocks/sec'
        $counters += '\SQLServer:Memory Manager\Memory Grants Pending'
        $counters += '\SQLServer:Wait Statistics(*)\Lock waits'
        $counters += '\SQLServer:Wait Statistics(*)\Log buffer waits'
        $counters += '\SQLServer:Wait Statistics(*)\Log write waits'
        $counters += '\SQLServer:Wait Statistics(*)\Network IO waits'
        $counters += '\SQLServer:Wait Statistics(*)\Non-Page latch waits'
        $counters += '\SQLServer:Wait Statistics(*)\Page IO latch waits'
        $counters += '\SQLServer:Wait Statistics(*)\Page latch waits'

    }

    $parameters = @{ Counter = $counters; SampleInterval = $SampleInterval }

    if ( $ComputerName ) {
        $parameters['ComputerName'] = $ComputerName
    }

    if ( $MaxSamples ) {
        $parameters['MaxSamples'] = $MaxSamples
    }

    if ( $Continuous ) {
        $parameters['Continuous'] = $Continuous
    }

    Get-Counter @parameters

}


<#
.SYNOPSIS
   Удаление неиспользуемых объектов конфигурации

.DESCRIPTION
   Удаление элементов конфигурации с синонимом "(не используется)"

.EXAMPLE
   PS C:\> $modules = Remove-1CNotUsedObjects E:\TEMP\ExportingConfiguration
   PS C:\> $gr = $modules | group File, Object | select -First 1
   PS C:\> ise ($gr.Group.File | select -First 1) # открываем модуль в новой вкладке ISE
   # альтернатива 'start notepad $gr.Group.File[0]'
   PS C:\> $gr.Group | select Object, Type, Line, Position -Unique | sort Line, Position | fl # Смотрим что корректировать
   PS C:\>  $modules = $modules | ? File -NE ($gr.Group.File | select -First 1) # удаление обработанного файла из списка объектов
   # альтернатива '$modules = $modules | ? File -NE $psise.CurrentFile.FullPath'
   # и все сначала с команды '$gr = $modules | group File, Object | select -First 1'

.NOTES  
    Name: 1CHelper
    Author: yauhen.makei@gmail.com

.LINK  
    https://github.com/emakei/1CHelper.psm1

.INPUTS
   Пусть к файлам выгрузки конфигурации

.OUTPUTS
   Массив объектов с описанием файлов модулей и позиций, содержащих упоминания удаляемых объектов

#>
function Remove-1CNotUsedObjects
{
    [CmdletBinding(DefaultParameterSetName='pathToConfigurationFiles', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='Medium')]
    [OutputType([Object[]])]
    Param
    (
        # Путь к файлам выгрузки конфигурации
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   HelpMessage='Путь к файлам выгрузки конфигурации')]
        [ValidateNotNullOrEmpty()]
        [Alias("pathToFiles")] 
        [ValidateScript({Test-Path -LiteralPath $_})]
        [string]
        $pathToConfigurationFiles
    )

    Begin
    {
        Write-Verbose "Начало обработки файлов в $(Get-Date)"
        # Переменная для поиска подстроки определения типа платформы в строке с типом метаданных
        $chars = [char[]]("A")
        66..90 | ForEach-Object { $chars += [char]$_ }
        # Содержит псевдонимы пространств имен для XPath
        $hashTable = @{ root = "http://v8.1c.ru/8.3/MDClasses"; core = "http://v8.1c.ru/8.1/data/core"; readable = "http://v8.1c.ru/8.3/xcf/readable" }
        # Используется для приведения строки в нижний регистр при поиске подстроки "(не используется)"
        $dict = @{ replace = "НЕИСПОЛЬЗУТЯ"; with = "неиспользутя" }
        # эти данные не требуется обрабатывать
        $excludeTypes = @('Template','Help','WSDefinition')
    }
    Process
    {
        if ($pscmdlet.ShouldProcess("Обработать файлы в каталоге '$pathToFiles'"))
        {
            # Содержит имена на удаляемых файлов
            $fileRefs = [string[]]("")
            # Содержит имена файлов модулей, в которых упоминаются не используемые объекты
            $modules = @()
            # Содержит имена типов объектов конфигурации для удаления
            $typeRefs = [string[]]("")
            # Содержит имена типов дочерних объектов конфигурации (формы, команды, реквизиты, ресурсы, макеты)
            $childRefs = [string[]]("")
            # Выборка файлов вида <ИмяТипаПлатформы>.<ИмяТипаМетаданных>.xml
            Write-Progress -Activity "Поиск файлов *.xml" -Completed 
            $files = Get-ChildItem -LiteralPath $pathToConfigurationFiles -Filter *.xml -File #| ? { $_.Name.ToString().Split('.').Count -eq 3 }
            $i = 1
            foreach ($item in $files) {
                Write-Progress -Activity "Поиск не используемых элементов в файлах *.xml" -PercentComplete ( $i / $files.Count * 100 )
                $thisIsShort = ($item.Name.ToString().Split('.').Count -eq 3)
                if ( $item.Name.Split('.').Count % 2 -eq 1 ) {
                    $pref = $item.Name.Split('.')[$item.Name.Split('.').Count-3]
                    $name = $item.Name.Split('.')[$item.Name.Split('.').Count-2]
                } else {
                    $pref = $item.Name.Split('.')[$item.Name.Split('.').Count-2]
                    $name = $item.Name.Split('.')[$item.Name.Split('.').Count-3]
                }
                $addAll = $false
                [xml]$xml = Get-Content $item.FullName -Encoding UTF8
                $nsmgr = New-Object -TypeName System.Xml.XmlNamespaceManager($xml.NameTable)
                # Добавляем псевдоним для 'v8'
                $nsmgr.AddNamespace("core", "http://v8.1c.ru/8.1/data/core")
                $nsmgr.AddNamespace("root", "http://v8.1c.ru/8.3/MDClasses")
                $nsmgr.AddNamespace("item", "http://v8.1c.ru/8.3/xcf/readable")
                $nsmgr.AddNamespace("exch", "http://v8.1c.ru/8.3/xcf/extrnprops")
                # Если синоним объекта содержит подстроку '(не используется)', то
                if ($thisIsShort -and ($xml | `
                    Select-Xml -XPath ("//root:$pref/root:Properties/root:Synonym/core:item/core:content[contains(translate(text(),$($dict.replace),$($dict.with)),'(не используется)')]") `
                        -Namespace $hashTable `
                        | Measure-Object).Count -gt 0) {
                    # Добавляем имя файла в массив удаляемых
                    $fileRefs += $item.FullName
                    # Находим производные типы
                    $tmp = $xml | Select-Xml -XPath ("//root:$pref/root:InternalInfo/readable:GeneratedType/@name") -Namespace $hashTable
                    $tmp | ForEach-Object { $typeRefs += $_.ToString() }
                    # Находим подчиненные объекты (<ИмяТипаПлатформы>.<ИмяТипаМетаданных>.*) и добавляем к удаляемым файлам
                    Get-ChildItem -LiteralPath $pathToFiles -Filter "$($m[0]).$($m[1]).*" -File | Where-Object { $_.Name.ToString().Split('.').Count -gt 3 } | ForEach-Object { $fileRefs += $_.FullName }
                    $addAll = $true
                } elseif(-not $thisIsShort) {
                    
                }
                # Поиск аттрибутов
                if ($addAll) {
                    # Поиск аттрибутов
                    $xml.SelectNodes("//root:MetaDataObject/root:$pref/root:ChildObjects/root:Attribute/root:Properties/root:Name", $nsmgr) | ForEach-Object { $childRefs += "$pref.$name.Attribute.$($_.'#text')" } 
                    # Поиск форм
                    $xml.SelectNodes("//root:MetaDataObject/root:$pref/root:ChildObjects/root:Form", $nsmgr) | ForEach-Object { $childRefs += "$pref.$name.Form.$($_.'#text')" }
                    # Поиск команд
                    $xml.SelectNodes("//root:MetaDataObject/root:$pref/root:ChildObjects/root:Command/root:Properties/root:Name", $nsmgr) | ForEach-Object { $childRefs += "$pref.$name.Command.$($_.'#text')" }
                    # Поиск макетов
                    $xml.SelectNodes("//root:MetaDataObject/root:$pref/root:ChildObjects/root:Template", $nsmgr) | ForEach-Object { "$pref.$name.Template.$($_.'#text')" }
                    # Поиск ресурсов информациооного регистра
                    if ($pref -eq 'InformationRegister') {
                        $xml.SelectNodes("//root:MetaDataObject/root:$pref/root:ChildObjects/root:Resource/root:Properties/root:Name", $nsmgr) | ForEach-Object { "$pref.$name.Resource.$($_.'#text')" }
                    }
                } else {
                    <# 
                    # Если синоним объекта содержит текст "(не используется)", тогда удаляем файл
                    if (($xml | `
                        Select-Xml -XPath ("//root:$pref/root:Properties/root:Synonym/core:item/core:content[contains(translate(text(),$($dict.replace),$($dict.with)),'(не используется)')]") `
                            -Namespace $hashTable `
                            | measure).Count -gt 0) {
                        # Удаление файлов
                        rm ($item.Name.Substring(0, $item.Name.Length - $item.Extension.Length) + '*') -Verbose
                    } 
                    #>
                }
                $i++
            }
            # Удаляем файлы
            $fileRefs | Where-Object { $_ -notlike '' } | ForEach-Object {Remove-Item $_ -Verbose}
            # Выбираем оставшиеся для поиска неиспользуемых ссылок на типы и атрибутов
            Write-Progress -Activity "Поиск файлов *.xml" -Completed -Status "Подготовка"
            $filesToUpdate = Get-ChildItem -LiteralPath $pathToConfigurationFiles -Filter *.xml -File
            # Удаляем пустой элемент (Создан при вызове конструктора типа)
            Write-Progress -Activity "Обработка ссылок для поиска" -Completed -Status "Подготовка"
            $typeRefs = $typeRefs | Where-Object { $_ -notlike '' } | Select-Object -Unique
            $childRefs = $childRefs | Where-Object { $_ -notlike '' } | Select-Object -Unique
            $i = 1
            foreach ( $item in $filesToUpdate ) {
                Write-Progress -Activity "Обработка файлов *.xml" -PercentComplete ( $i / $filesToUpdate.Count * 100 )
                Write-Verbose "Файл '$($item.FullName)'"
                if ( $item.Name.Split('.').Count % 2 -eq 1 ) {
                    $pref = $item.Name.Split('.')[$item.Name.Split('.').Count-3]
                    $name = $item.Name.Split('.')[$item.Name.Split('.').Count-2]
                } else {
                    $pref = $item.Name.Split('.')[$item.Name.Split('.').Count-2]
                    $name = $item.Name.Split('.')[$item.Name.Split('.').Count-3]
                }
                if ($pref -in $excludeTypes) { 
                    Write-Verbose "Пропуск файла по шаблону '$pref'"
                    Continue
                }
                [xml]$xml = Get-Content $item.FullName -Encoding UTF8
                # Создаем менеджер пространств имен для XPath
                $nsmgr = New-Object -TypeName System.Xml.XmlNamespaceManager($xml.NameTable)
                # Добавляем псевдоним для 'v8'
                $nsmgr.AddNamespace("core", "http://v8.1c.ru/8.1/data/core")
                $nsmgr.AddNamespace("root", "http://v8.1c.ru/8.3/MDClasses")
                $nsmgr.AddNamespace("item", "http://v8.1c.ru/8.3/xcf/readable")
                $nsmgr.AddNamespace("exch", "http://v8.1c.ru/8.3/xcf/extrnprops")
                # Если это файл описания конфигурации
                #else
                if ($item.Name -eq 'Configuration.xml') {
                    Write-Verbose "in 'Configuration'"
                    foreach ($tref in $typeRefs) {
                        # Получаем из <ИмяТипаПлатформы>.<ИмяТипаМетаданных> значение <ИмяТипаПлатформы>
                        if ($tref.ToString().Split('.').Count -lt 2) {
                            Write-Error "Неверный тип для поиска" -ErrorId C1 -Targetobject $tref -Category ParserError
                            continue
                        }
                        $tpref = $tref.ToString().Split('.')[0] 
                        $max = -1
                        $chars | ForEach-Object { $max = [Math]::Max($max, $tpref.LastIndexOf($_)) }
                        if ($max -eq -1) {
                            Write-Error "Неверный тип для поиска" -ErrorId C2 -Targetobject $tref -Category ParserError
                            continue
                        } else {
                            $type = if($max -eq 0) { $tref.Split('.')[0] } else { $tpref.Substring(0, $max) }
                            try {
                                $xml.SelectNodes("//root:MetaDataObject/root:Configuration/root:ChildObjects/root:$type[text()='$($tref.Split('.')[1])']/.", $nsmgr) `
                                    | ForEach-Object { $_.ParentNode.RemoveChild($_) | Out-Null }
                            } catch {
                                Write-Error "Ошибка обработки файла" -ErrorId C3 `
                                    -Targetobject "//root:MetaDataObject/root:Configuration/root:ChildObjects/root:$type[text()='$($tref.Split('.')[1])']/."
                                    -Category ParserError
                            }
                        }
                    }
                }
                # Если это файл описания командного интерфейса конфигурации
                elseif ($item.Name -eq 'Configuration.CommandInterface.xml') {
                    Write-Verbose "in 'Configuration.CommandInterface'"
                    foreach ($tref in $typeRefs) {
                        # Обрабатываем только роли и подсистемы
                        if (-not ($tref.StartsWith('Role.') -or $tref.StartsWith('Subsystem.'))) { Continue }
                        # Получаем из <ИмяТипаПлатформы>.<ИмяТипаМетаданных> значение <ИмяТипаПлатформы>
                        if ($tref.ToString().Split('.').Count -lt 2) {
                            Write-Error "Неверный тип для поиска" -ErrorId C1 -Targetobject $tref -Category ParserError
                            continue
                        }
                        $tpref = $tref.ToString().Split('.')[0] 
                        $max = -1
                        $chars | ForEach-Object { $max = [Math]::Max($max, $tpref.LastIndexOf($_)) }
                        if ($max -eq -1) {
                            Write-Error "Неверный тип для поиска" -ErrorId C2 -Targetobject $tref -Category ParserError
                            continue
                        } else {
                            $type = if($max -eq 0) { $tref.Split('.')[0] } else { $tpref.Substring(0, $max) }
                            if ($tref.StartsWith('Role.')) {
                                try {
                                    $xml.SelectNodes("//exch:$name/exch:SubsystemsVisibility/exch:Subsystem/exch:Visibility/item:Value[@name='$tref']/.", $nsmgr) `
                                        | ForEach-Object { $_.ParentNode.RemoveChild($_) | Out-Null }
                                } catch {
                                    Write-Error "Ошибка обработки файла" -ErrorId C3 `
                                        -Targetobject "//exch:$name/exch:SubsystemsVisibility/exch:Subsystem/exch:Visibility/item:Value[@name='$tref']/."
                                        -Category ParserError
                                }
                            } else {
                                try {
                                    $xml.SelectNodes("//exch:$name/exch:SubsystemsVisibility/exch:Subsystem[@name='$tref']/.", $nsmgr) `
                                        | ForEach-Object { $_.ParentNode.RemoveChild($_) | Out-Null }
                                } catch {
                                    Write-Error "Ошибка обработки файла" -ErrorId C3 `
                                        -Targetobject "//exch:$name/exch:SubsystemsVisibility/exch:Subsystem[@name='$tref']/."
                                        -Category ParserError
                                }
                                try {
                                    $xml.SelectNodes("//exch:$name/exch:SubsystemsOrder/exch:Subsystem[text()='$tref']/.", $nsmgr) `
                                        | ForEach-Object { $_.ParentNode.RemoveChild($_) | Out-Null }
                                } catch {
                                    Write-Error "Ошибка обработки файла" -ErrorId C3 `
                                        -Targetobject "//exch:$name/exch:SubsystemsOrder/exch:Subsystem[text()='$tref']/."
                                        -Category ParserError
                                }
                            }
                        }
                    }
                }
                # Если это файл описания командного интерфейса подсистемы
                elseif ($pref -eq 'CommandInterface') {
                    <#Write-Verbose "in 'Subsystem.*.CommandInterface'"
                    foreach ($tref in $typeRefs) {
                        # Обрабатываем только роли и подсистемы
                        if (-not ($tref.StartsWith('Role.') -or $tref.StartsWith('Subsystem.'))) { Continue }
                        # Получаем из <ИмяТипаПлатформы>.<ИмяТипаМетаданных> значение <ИмяТипаПлатформы>
                        if ($tref.ToString().Split('.').Count -lt 2) {
                            Write-Error "Неверный тип для поиска" -ErrorId C1 -Targetobject $tref -Category ParserError
                            continue
                        }
                        $tpref = $tref.ToString().Split('.')[0] 
                        $max = -1
                        $chars | % { $max = [Math]::Max($max, $tpref.LastIndexOf($_)) }
                        if ($max -eq -1) {
                            Write-Error "Неверный тип для поиска" -ErrorId C2 -Targetobject $tref -Category ParserError
                            continue
                        } else {
                            $type = if($max -eq 0) { $tref.Split('.')[0] } else { $tpref.Substring(0, $max) }
                            if ($tref.StartsWith('Role.')) {
                                try {
                                    $xml.SelectNodes("//exch:$name/exch:SubsystemsVisibility/exch:Subsystem/exch:Visibility/item:Value[@name='$tref']/.", $nsmgr) `
                                        | % { $_.ParentNode.RemoveChild($_) | Out-Null }
                                } catch {
                                    Write-Error "Ошибка обработки файла" -ErrorId C3 `
                                        -Targetobject "//exch:$name/exch:SubsystemsVisibility/exch:Subsystem/exch:Visibility/item:Value[@name='$tref']/."
                                        -Category ParserError
                                }
                            } else {
                                try {
                                    $xml.SelectNodes("//exch:$name/exch:SubsystemsVisibility/exch:Subsystem[@name='$tref']/.", $nsmgr) `
                                        | % { $_.ParentNode.RemoveChild($_) | Out-Null }
                                } catch {
                                    Write-Error "Ошибка обработки файла" -ErrorId C3 `
                                        -Targetobject "//exch:$name/exch:SubsystemsVisibility/exch:Subsystem[@name='$tref']/."
                                        -Category ParserError
                                }
                                try {
                                    $xml.SelectNodes("//exch:$name/exch:SubsystemsOrder/exch:Subsystem[text()='$tref']/.", $nsmgr) `
                                        | % { $_.ParentNode.RemoveChild($_) | Out-Null }
                                } catch {
                                    Write-Error "Ошибка обработки файла" -ErrorId C3 `
                                        -Targetobject "//exch:$name/exch:SubsystemsOrder/exch:Subsystem[text()='$tref']/."
                                        -Category ParserError
                                }
                            }
                        }
                    }#>
                }
                # Если это файл описания подсистемы
                elseif (($pref -eq 'Subsystem') -and ($item.Name.Split('.').Count % 2 -eq 1)) {
                    Write-Verbose "in 'Subsystem'"
                    foreach ($tref in $typeRefs) {
                        # Получаем из <ИмяТипаПлатформы>.<ИмяТипаМетаданных> значение <ИмяТипаПлатформы>
                        if ($tref.ToString().Split('.').Count -lt 2) {
                            Write-Error "Неверный тип для поиска" -ErrorId S1 -Targetobject $tref -Category ParserError
                            continue
                        }
                        $tpref = $tref.ToString().Split('.')[0]
                        $max = -1
                        $chars | ForEach-Object { $max = [Math]::Max($max, $tpref.LastIndexOf($_)) }
                        if ($max -eq -1) {
                            Write-Error "Неверный тип для поиска" -ErrorId S2 -Targetobject $tref -Category ParserError
                            continue
                        } else {
                            $type = if($max -eq 0) { $tref.Split('.')[0] } else { $tpref.Substring(0, $max) }
                            try {
                                $xml.SelectNodes("//root:MetaDataObject/root:Subsystem/root:Properties/root:Content/item:Item[text()='$($type+'.'+$tref.Split('.')[1])']/.", $nsmgr) `
                                    | ForEach-Object { $_.ParentNode.RemoveChild($_) | Out-Null }
                            } catch {
                                Write-Error "Ошибка обработки файла" -ErrorId S3 `
                                    -Targetobject "//root:MetaDataObject/root:Subsystem/root:Properties/root:Content/item:Item[text()='$($type+'.'+$tref.Split('.')[1])']/."
                                    -Category ParserError
                            }
                        }
                    }
                }
                # Если это файл описания состава плана обмена
                elseif ($pref -eq 'Content') {
                    Write-Verbose "in 'Content'"
                    foreach ($tref in $typeRefs) {
                        # Получаем из <ИмяТипаПлатформы>.<ИмяТипаМетаданных> значение <ИмяТипаПлатформы>
                        if ($tref.ToString().Split('.').Count -lt 2) {
                            Write-Error "Неверный тип для поиска" -ErrorId S1 -Targetobject $tref -Category ParserError
                            continue
                        }
                        $tpref = $tref.ToString().Split('.')[0]
                        $max = -1
                        $chars | ForEach-Object { $max = [Math]::Max($max, $tpref.LastIndexOf($_)) }
                        if ($max -eq -1) {
                            Write-Error "Неверный тип для поиска" -ErrorId S2 -Targetobject $tref -Category ParserError
                            continue
                        } else {
                            $type = if($max -eq 0) { $tref.Split('.')[0] } else { $tpref.Substring(0, $max) }
                            try {
                                $xml.SelectNodes("//exch:$name/exch:Item/exch:Metadata/[text()='$($type+'.'+$tref.Split('.')[1])']/.", $nsmgr) `
                                    | ForEach-Object { $_.ParentNode.RemoveChild($_) | Out-Null }
                            } catch {
                                Write-Error "Ошибка обработки файла" -ErrorId S3 `
                                    -Targetobject "//exch:$name/exch:Item/exch:Metadata/[text()='$($type+'.'+$tref.Split('.')[1])']/."
                                    -Category ParserError
                            }
                        }
                    }
                }
                # Если это файл описания прав доступа
                elseif ($pref -eq 'Rights') {
                }
                # Иначе удаляем ссылки на неиспользуемые типы и узлы с синонимом содержащим текст "(не используется)" 
                else {
                    Write-Verbose "Поиск ссылок" 
                    $typeRefs | ForEach-Object { $xml.SelectNodes("//*/core:Type[contains(text(), '$_')]/.", $nsmgr) } | ForEach-Object { $_.ParentNode.RemoveChild($_) | Out-Null }
                    Write-Verbose "Поиск неиспользуемых атрибутов"
                    $xml.SelectNodes("//*/core:content[contains(translate(text(),$($dict.replace),$($dict.with)),'(не используется)')]/../../../..", $nsmgr) `
                        | ForEach-Object { $_.ParentNode.RemoveChild($_) | Out-Null }
                }
                if (Test-Path -LiteralPath $item.FullName) {
                    $xml.Save($item.FullName)
                }
                $i++
            }
            # Обработка модулей объектов
            Write-Progress -Activity "Поиск файлов модулей (*.txt)" -Completed
            $txtFiles = Get-ChildItem -LiteralPath $pathToConfigurationFiles -Filter *.txt -File
            $i = 1
            foreach ( $item in $txtFiles ) {
                Write-Progress -Activity "Обработка файлов *.txt" -PercentComplete ( $i / $txtFiles.Count * 100 )
                Write-Verbose "Файл '$($item.FullName)'"
                $data = Get-Content $item.FullName -Encoding UTF8
                $lineNumber = 0
                foreach ( $str in $data ) {
                    $lineNumber += 1
                    # Если строка закомментирована - продолжаем
                    if ($str -match "\A[\t, ]*//") { continue }
                    foreach ( $tref in $typeRefs ) {
                        try {
                            $subString = $tref.ToString().Split('.')[1]
                            $subStringType = $tref.ToString().Split('.')[0]
                        } catch {
                            Write-Error "Неверный тип для поиска" -ErrorId T1 -Targetobject $tref -Category ParserError
                            continue
                        }
                        $ind = $str.IndexOf($subString)
                        if ($ind -ne -1) {
                            # Костыль
                            $modules += 1 | Select-Object @{ Name = 'File';     Expression = { $item.FullName } },
                                                          @{ Name = 'Line';     Expression = { $lineNumber } },
                                                          @{ Name = 'Position'; Expression = { $ind + 1 } },
                                                          @{ Name = 'Object';   Expression = { $subString } },
                                                          @{ Name = 'Type';     Expression = { $subStringType } }
                            Write-Verbose "`$tref = $tref; `$lineNumber = $lineNumber; `$ind = $ind`n`$subString = '$subString'"
                        }
                    }
                }
                $i++
            }
            Write-Output $modules
        }
    }
    End
    {
        Write-Verbose "Окончание обработки файлов в $(Get-Date)"
    }
}

<#
.SYNOPSIS
   Поиск стартера 1С

.DESCRIPTION
   Поиск исполняемого файла 1cestart.exe

.NOTES  
    Name: 1CHelper
    Author: yauhen.makei@gmail.com

.LINK  
    https://github.com/emakei/1CHelper.psm1

.EXAMPLE
   Find-1CEstart

.OUTPUTS
   NULL или строку с полным путём к исполняемому файлу
#>

function Find-1CEstart
{
    Param(
        [CmdletBinding()]
        # Имя компьютера для поиска версии
        [string]$ComputerName = $env:COMPUTERNAME
    )
    
    $pathToStarter = $null

    $keys = @( @{ leaf='ClassesRoot'; path='Applications\\1cestart.exe\\shell\\open\\command' } )
    $keys += @{ leaf='ClassesRoot'; path='V83.InfoBaseList\\shell\\open\\command' }
    $keys += @{ leaf='ClassesRoot'; path='V83.InfoBaseListLink\\shell\\open\\command' }
    $keys += @{ leaf='ClassesRoot'; path='V82.InfoBaseList\\shell\\open\\command' }
    $keys += @{ leaf='LocalMachine'; path='SOFTWARE\\Classes\\Applications\\1cestart.exe\\shell\\open\\command' }
    $keys += @{ leaf='LocalMachine'; path='SOFTWARE\\Classes\\V83.InfoBaseList\\shell\\open\\command' }
    $keys += @{ leaf='LocalMachine'; path='SOFTWARE\\Classes\\V83.InfoBaseListLink\\shell\\open\\command' }
    $keys += @{ leaf='LocalMachine'; path='SOFTWARE\\Classes\\V82.InfoBaseList\\shell\\open\\command' }

    foreach( $key in $keys ) {
                
         Try {
             $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey( $key.leaf, $computerName )
         } Catch {
             Write-Error $_
             Continue
         }
 
         $regkey = $reg.OpenSubKey( $key.path )

         If( -not $regkey ) {
            $index = 0
         } else {
            $defaultValue = $regkey.GetValue("").ToString()
            $index = $defaultValue.IndexOf("1cestart.exe")
         }
         

         if ( $index -gt 0 ) {

            if ( $defaultValue[0] -eq '"' ) {
                $pathToStarter = $defaultValue.Substring( 1, $index + 11 )
            } else {
                $pathToStarter = $defaultValue.Substring( 0, $index + 11 )
            }

            $reg.Close()
            Break

         }

         $reg.Close()

    }

    # если не удалось найти, то пробуем через WinRM

    if ( -not $pathToStarter ) {
        
        $scriptBlock = {
            if (Test-Path "${env:ProgramFiles}\1cv8\common\1cestart.exe" ) {
                "${env:ProgramFiles}\1cv8\common\1cestart.exe"
            }
            elseif ( Test-Path "${env:ProgramFiles(x86)}\1cv8\common\1cestart.exe" ) {
                "${env:ProgramFiles(x86)}\1cv8\common\1cestart.exe" 
            }
            elseif ( Test-Path "${env:ProgramFiles(x86)}\1cv82\common\1cestart.exe" ) {
                "${env:ProgramFiles(x86)}\1cv82\common\1cestart.exe"
            }
            elseif ( Test-Path "$ENV:USERPROFILE\AppData\Local\Programs\1cv8" ) {
                "$ENV:USERPROFILE\AppData\Local\Programs\1cv8\common\1cestart.exe"
            }
            elseif ( Test-Path "$ENV:USERPROFILE\AppData\Local\Programs\1cv8_x86" ) {
                "$ENV:USERPROFILE\AppData\Local\Programs\1cv8_x86\common\1cestart.exe"
            }
            else { $null }
        }

        $pathToStarter = if ( $ComputerName -ne $env:COMPUTERNAME ) { Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ErrorAction Continue} else { $scriptBlock.Invoke() }
    }

    $pathToStarter
}

<#
.SYNOPSIS
   Поиск строк подключения 1С

.DESCRIPTION
   Поиск строк подключения

.NOTES  
    Name: 1CHelper
    Author: yauhen.makei@gmail.com

.LINK  
    https://github.com/emakei/1CHelper.psm1

.EXAMPLE
    Find-1C8conn -NoCommonFiles -UseFilesFromDirectories ("$ENV:USERPROFILE\Desktop","D:\") | select Line

.EXAMPLE
    Find-1C8conn | select Line

.EXAMPLE
    Find-1C8conn -ReturnValue Files

.EXAMPLE
    Find-1C8conn -UseFilesFromDirectories ("$ENV:USERPROFILE\Desktop","D:\")

.EXAMPLE
    Find-1C8conn -UseFilesFromDirectories ("$ENV:USERPROFILE\Desktop","D:\")

.EXAMPLE
    Find-1C8conn -NoCommonFiles -UseFilesFromDirectories ("$ENV:USERPROFILE\Desktop","D:\") -ReturnValue Files

.OUTPUTS
    Список совпадений по шаблону строки подключения в файлах

.OUTPUTS
    Список найденных файлов *.v8i

#>
function Find-1C8conn
{
    [OutputType([Object[]])]
    Param(
        # Использовать общие файлы
        [switch]$NoCommonFiles,
        [string[]]$UseFilesFromDirectories,
        [ValidateSet('ConnectionStrings','Files')]
        [string]$ReturnValue = 'ConnectionStrings'
    )

    $commonIbFiles = @()

    if ($UseFilesFromDirectories)
    {
        $commonIbFiles += Get-ChildItem $UseFilesFromDirectories -File -Filter *.v8i
    }
    
    if (-not $NoCommonFiles -and (Test-Path "$ENV:APPDATA\1C\1CEStart\ibases.v8i"))
    {
        $commonIbFiles += Get-Item "$ENV:APPDATA\1C\1CEStart\ibases.v8i"
    }

    if (-not $NoCommonFiles -and (Test-Path "$ENV:ALLUSERSPROFILE\1C\1CEStart\ibases.v8i"))
    {
        $commonIbFiles += Get-Item "$ENV:ALLUSERSPROFILE\1C\1CEStart\ibases.v8i"
    }
    
    $list = @()

    switch( $ReturnValue )
    {
        'ConnectionStrings'
        {
            $list += $commonIbFiles | Select-String -Pattern '\s*Connect\s*\=.*$'
        }
        'Files'
        {
            $list += $commonIbFiles
        }
    }

    $list
    
}

<#
.SYNOPSIS
    Собирает информацию с кластеров 1С

.DESCRIPTION

.NOTES  
    Name: 1CHelper
    Author: yauhen.makei@gmail.com

.LINK  
    https://github.com/emakei/1CHelper.psm1

.EXAMPLE
    Get-1ClusterData

.EXAMPLE
    Get-1ClusterData 'srv-01','srv-02'

.EXAMPLE
    $netHaspParams = Get-1CNetHaspIniStrings
    $hostsToQuery += $netHaspParams.NH_SERVER_ADDR
    $hostsToQuery += $netHaspParams.NH_SERVER_NAME
    
    $stat = $hostsToQuery | % { Get-1CclusterData $_ -Verbose }

.OUTPUTS
    Данные кластера
#>
function Get-1ClusterData
{
[OutputType([Object[]])]
[CmdletBinding()]
Param(
    # Адрес хоста для сбора статистики
    [Parameter(Mandatory=$true)]
    [string]$HostName,
    # имя админитратора кластера
    [Parameter(Mandatory=$true)]
    [string]$User,
    # пароль администратора кластера
    [Parameter(Mandatory=$true)]
    [Security.SecureString]$Password,
    # не получать инфорацию об администраторах кластера
    [switch]$NoClusterAdmins=$false,
    # не получать инфорацию о менеджерах кластера
    [switch]$NoClusterManagers=$false,
    # не получать инфорацию о рабочих серверах
    [switch]$NoWorkingServers=$false,
    # не получать инфорацию о рабочих процессах
    [switch]$NoWorkingProcesses=$false,
    # не получать инфорацию о сервисах кластера
    [switch]$NoClusterServices=$false,
    # Получать информацию о соединениях только для кластера, везде или вообще не получать
    [ValidateSet('None', 'Cluster', 'Everywhere')]
    [string]$ShowConnections='Everywhere',
    # Получать информацию о сессиях только для кластера, везде или вообще не получать
    [ValidateSet('None', 'Cluster', 'Everywhere')]
    [string]$ShowSessions='Everywhere',
    # Получать информацию о блокировках только для кластера, везде или вообще не получать
    [ValidateSet('None', 'Cluster', 'Everywhere')]
    [string]$ShowLocks='Everywhere',
    # не получать инфорацию об информационных базах
    [switch]$NoInfobases=$false,
    # не получать инфорацию о требованиях назначения
    [switch]$NoAssignmentRules=$false,
    # верия компоненты
    [ValidateSet(2, 3)]
    [int]$Version=3
    )

Begin {
    $connector = New-Object -ComObject "v8$version.COMConnector"
    }

Process {
              
    $obj = 1 | Select-Object  @{ name = 'Host';     Expression = { $HostName } }`
                            , @{ name = 'Error';    Expression = { '' } }`
                            , @{ name = 'Clusters'; Expression = {  @() } }

    try {
        Write-Verbose "Подключение к '$HostName'"
        $connection = $connector.ConnectAgent( $HostName )
        $abort = $false
    } catch {
        Write-Warning $_
        $obj.Error = $_.Exception.Message
        $result = $obj
        $abort = $true
    }
        
    if ( -not $abort ) {
            
        Write-Verbose "Подключен к `"$($connection.ConnectionString)`""

        $clusters = $connection.GetClusters()

        foreach( $cluster in $clusters ) {
                
            $cls = 1 | Select-Object  @{ name = 'ClusterName';                Expression = { $cluster.ClusterName } }`
                                    , @{ name = 'ExpirationTimeout';          Expression = { $cluster.ExpirationTimeout } }`
                                    , @{ name = 'HostName';                   Expression = { $cluster.HostName } }`
                                    , @{ name = 'LoadBalancingMode';          Expression = { $cluster.LoadBalancingMode } }`
                                    , @{ name = 'MainPort';                   Expression = { $cluster.MainPort } }`
                                    , @{ name = 'MaxMemorySize';              Expression = { $cluster.MaxMemorySize } }`
                                    , @{ name = 'MaxMemoryTimeLimit';         Expression = { $cluster.MaxMemoryTimeLimit } }`
                                    , @{ name = 'SecurityLevel';              Expression = { $cluster.SecurityLevel } }`
                                    , @{ name = 'SessionFaultToleranceLevel'; Expression = { $cluster.SessionFaultToleranceLevel } }`
                                    , @{ name = 'Error';                      Expression = {} }

            Write-Verbose "Получение информации кластера `"$($cluster.ClusterName)`" на `"$($cluster.HostName)`""
                
            try {
                Write-Verbose "Аутентификация в кластере $($cluster.HostName,':',$cluster.MainPort,' - ',$cluster.ClusterName)"
                $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
                $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                $connection.Authenticate( $cluster, $User, $PlainPassword )
                $abort = $false
            } catch {
                Write-Warning $_
                $cls.Error = $_.Exception.Message
                $obj.Clusters += $cls
                $result = $obj
                $abort = $true
            }

            if ( -not $abort ) {
                    
                # TODO возможно нужно получить информацию из 'GetAgentAdmins'

                if ( -not $NoClusterAdmins ) {
                   
                    $admins = $connection.GetClusterAdmins( $cluster )
                    $objAdmin = @()

                    foreach ( $admin in $admins ) {
                        $objAdmin += 1 | Select-Object  @{ Name = 'Name';                Expression = { $admin.Name } },
                                                        @{ Name = 'Descr';               Expression = { $admin.Descr } },
                                                        @{ Name = 'PasswordAuthAllowed'; Expression = { $admin.PasswordAuthAllowed } },
                                                        @{ Name = 'SysAuthAllowed';      Expression = { $admin.SysAuthAllowed } },
                                                        @{ Name = 'SysUserName';         Expression = { $admin.SysUserName } }
                    }

                    Add-Member -InputObject $cls -Name ClusterAdmins -Value $objAdmin -MemberType NoteProperty
                    
                }

                if ( -not $NoClusterManagers ) {

                    $mngrs = $connection.GetClusterManagers( $cluster )
                    $objMngr = @()

                    foreach ( $mngr in $mngrs ) {
                        $objMngr += 1 | Select-Object  @{ Name = 'HostName';    Expression = { $mngr.HostName } },
                                                        @{ Name = 'Descr';       Expression = { $mngr.Descr } },
                                                        @{ Name = 'MainManager'; Expression = { $mngr.MainManager } },
                                                        @{ Name = 'MainPort';    Expression = { $mngr.MainPort } },
                                                        @{ Name = 'PID';         Expression = { $mngr.PID } }
                    }

                    Add-Member -InputObject $cls -Name ClusterManagers -Value $objMngr -MemberType NoteProperty

                }

                if ( -not $NoWorkingServers ) {

                    $ws = $connection.GetWorkingServers( $cluster )
                    $objWS = @()
                    foreach( $workingServer in $ws ) {

                        $objWS += 1 | Select-Object @{ Name = 'ClusterMainPort';                   Expression = { $workingServer.ClusterMainPort } },
                                                    @{ Name = 'ConnectionsPerWorkingProcessLimit'; Expression = { $workingServer.ConnectionsPerWorkingProcessLimit } },
                                                    @{ Name = 'DedicatedManagers';                 Expression = { $workingServer.DedicatedManagers } },
                                                    @{ Name = 'HostName';                          Expression = { $workingServer.HostName } },
                                                    @{ Name = 'InfoBasesPerWorkingProcessLimit';   Expression = { $workingServer.InfoBasesPerWorkingProcessLimit } },
                                                    @{ Name = 'MainPort';                          Expression = { $workingServer.MainPort } },
                                                    @{ Name = 'MainServer';                        Expression = { $workingServer.MainServer } },
                                                    @{ Name = 'Name';                              Expression = { $workingServer.Name } },
                                                    @{ Name = 'SafeCallMemoryLimit';               Expression = { $workingServer.SafeCallMemoryLimit } },
                                                    @{ Name = 'SafeWorkingProcessesMemoryLimit';   Expression = { $workingServer.SafeWorkingProcessesMemoryLimit } },
                                                    @{ Name = 'WorkingProcessMemoryLimit';         Expression = { $workingServer.WorkingProcessMemoryLimit } }

                        if ( -not $NoAssignmentRules ) {
                            
                            $assignmentRules = $connection.GetAssignmentRules( $cluster, $workingServer )
                            $objAR = @()
                            foreach( $assignmentRule in $assignmentRules ) {
                                $objAR += 1 | Select-Object @{ Name = 'ApplicationExt'; Expression = { $assignmentRule.ApplicationExt } },
                                                            @{ Name = 'InfoBaseName';   Expression = { $assignmentRule.InfoBaseName } },
                                                            @{ Name = 'ObjectType';     Expression = { $assignmentRule.ObjectType } },
                                                            @{ Name = 'Priority';       Expression = { $assignmentRule.Priority } },
                                                            @{ Name = 'RuleType';       Expression = { $assignmentRule.RuleType } }
                            }
                
                            Add-Member -InputObject $objWS[$objWS.Count-1] -Name AssignmentRules -Value $objAR -MemberType NoteProperty

                        }

                    }

                    Add-Member -InputObject $cls -Name WorkingServers -Value $objWS -MemberType NoteProperty

                }

                if ( -not $NoWorkingProcesses ) {
                
                    $wp = $connection.GetWorkingProcesses( $cluster )
                    $objWP = @()

                    foreach( $workingProcess in $wp ) {
                    
                        $objWP += 1 | Select-Object @{ Name = 'AvailablePerfomance'; Expression = { $workingProcess.AvailablePerfomance } },
                                                    @{ Name = 'AvgBackCallTime';     Expression = { $workingProcess.AvgBackCallTime } },
                                                    @{ Name = 'AvgCallTime';         Expression = { $workingProcess.AvgCallTime } },
                                                    @{ Name = 'AvgDBCallTime';       Expression = { $workingProcess.AvgDBCallTime } },
                                                    @{ Name = 'AvgLockCallTime';     Expression = { $workingProcess.AvgLockCallTime } },
                                                    @{ Name = 'AvgServerCallTime';   Expression = { $workingProcess.AvgServerCallTime } },
                                                    @{ Name = 'AvgThreads';          Expression = { $workingProcess.AvgThreads } },
                                                    @{ Name = 'Capacity';            Expression = { $workingProcess.Capacity } },
                                                    @{ Name = 'Connections';         Expression = { $workingProcess.Connections } },
                                                    @{ Name = 'HostName';            Expression = { $workingProcess.HostName } },
                                                    @{ Name = 'IsEnable';            Expression = { $workingProcess.IsEnable } },
                                                    @{ Name = 'License';             Expression = { try { $workingProcess.License.FullPresentation } catch { $null } } },
                                                    @{ Name = 'MainPort';            Expression = { $workingProcess.MainPort   } },
                                                    @{ Name = 'MemoryExcessTime';    Expression = { $workingProcess.MemoryExcessTime } },
                                                    @{ Name = 'MemorySize';          Expression = { $workingProcess.MemorySize } },
                                                    @{ Name = 'PID';                 Expression = { $workingProcess.PID } },
                                                    @{ Name = 'Running';             Expression = { $workingProcess.Running } },
                                                    @{ Name = 'SelectionSize';       Expression = { $workingProcess.SelectionSize } },
                                                    @{ Name = 'StartedAt';           Expression = { $workingProcess.StartedAt } },
                                                    @{ Name = 'Use';                 Expression = { $workingProcess.Use } }

                    }

                    Add-Member -InputObject $cls -Name WorkingProcesses -Value $objWP -MemberType NoteProperty

                }

                if ( -not $NoClusterServices ) {

                    $сs = $connection.GetClusterServices( $cluster )
                    $objCS = @()
                    foreach( $service in $сs ) {
                        $objCS += 1 | Select-Object @{ Name = 'Descr';    Expression = { $service.Descr } },
                                                    @{ Name = 'MainOnly'; Expression = { $service.MainOnly } },
                                                    @{ Name = 'Name';     Expression = { $service.Name } }
                        $objCM = @()
                        foreach( $cmngr in $service.ClusterManagers ) {
                            $objCM += 1 | Select-Object @{ Name = 'HostName';    Expression = { $cmngr.HostName } },
                                                        @{ Name = 'Descr';       Expression = { $cmngr.Descr } },
                                                        @{ Name = 'MainManager'; Expression = { $cmngr.MainManager } },
                                                        @{ Name = 'MainPort';    Expression = { $cmngr.MainPort } },
                                                        @{ Name = 'PID';         Expression = { $cmngr.PID } }
                        }
                        Add-Member -InputObject $objCS -Name ClusterManagers -Value $objCM -MemberType NoteProperty
                    }             

                }

                if ( $ShowConnections -ne 'None' ) {

                    $cConnections = $connection.GetConnections( $cluster )
                    $objCC = @()
                    foreach( $conn in $cConnections ) {
                        
                        $objCC += 1 | Select-Object @{ Name = 'Application'; Expression = { $conn.Application } },
                                                    @{ Name = 'blockedByLS'; Expression = { $conn.blockedByLS } },
                                                    @{ Name = 'ConnectedAt'; Expression = { $conn.ConnectedAt } },
                                                    @{ Name = 'ConnID';      Expression = { $conn.ConnID } },
                                                    @{ Name = 'Host';        Expression = { $conn.Host } },
                                                    @{ Name = 'InfoBase';    Expression = { @{ Descr = $conn.InfoBase.Descr; Name = $conn.InfoBase.Name } } },
                                                    @{ Name = 'SessionID';   Expression = { $conn.SessionID } },
                                                    @{ Name = 'Process';     Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object   @{ Name = 'AvailablePerfomance'; Expression = { $conn.Process.AvailablePerfomance } },
                                                                                                                @{ Name = 'AvgBackCallTime';     Expression = { $conn.Process.AvgBackCallTime } },
                                                                                                                @{ Name = 'AvgCallTime';         Expression = { $conn.Process.AvgCallTime } },
                                                                                                                @{ Name = 'AvgDBCallTime';       Expression = { $conn.Process.AvgDBCallTime } },
                                                                                                                @{ Name = 'AvgLockCallTime';     Expression = { $conn.Process.AvgLockCallTime } },
                                                                                                                @{ Name = 'AvgServerCallTime';   Expression = { $conn.Process.AvgServerCallTime } },
                                                                                                                @{ Name = 'AvgThreads';          Expression = { $conn.Process.AvgThreads } },
                                                                                                                @{ Name = 'Capacity';            Expression = { $conn.Process.Capacity } },
                                                                                                                @{ Name = 'Connections';         Expression = { $conn.Process.Connections } },
                                                                                                                @{ Name = 'HostName';            Expression = { $conn.Process.HostName } },
                                                                                                                @{ Name = 'IsEnable';            Expression = { $conn.Process.IsEnable } },
                                                                                                                @{ Name = 'License';             Expression = { try { $conn.Process.License.FullPresentation } catch { $null } } },
                                                                                                                @{ Name = 'MainPort';            Expression = { $conn.Process.MainPort   } },
                                                                                                                @{ Name = 'MemoryExcessTime';    Expression = { $conn.Process.MemoryExcessTime } },
                                                                                                                @{ Name = 'MemorySize';          Expression = { $conn.Process.MemorySize } },
                                                                                                                @{ Name = 'PID';                 Expression = { $conn.Process.PID } },
                                                                                                                @{ Name = 'Running';             Expression = { $conn.Process.Running } },
                                                                                                                @{ Name = 'SelectionSize';       Expression = { $conn.Process.SelectionSize } },
                                                                                                                @{ Name = 'StartedAt';           Expression = { $conn.Process.Process.StartedAt } },
                                                                                                                @{ Name = 'Use';                 Expression = { $conn.Process.Use } } } } }

                        if ( $ShowLocks -eq 'Everywhere' ) {

                            $locks = $connection.GetConnectionLocks( $cluster, $conn )
                            $objLock = @()
                            foreach( $lock in $locks ) {
                                $objLock += 1 | Select-Object @{ Name = 'Connection';  Expression = { if ( $lock.Connection ) { 1 | Select-Object @{ Name = 'Application'; Expression = { $lock.Connection.Application } },
                                                                                                                    @{ Name = 'blockedByLS'; Expression = { $lock.Connection.blockedByLS } },
                                                                                                                    @{ Name = 'ConnectedAt'; Expression = { $lock.Connection.ConnectedAt } },
                                                                                                                    @{ Name = 'ConnID';      Expression = { $lock.Connection.ConnID } },
                                                                                                                    @{ Name = 'Host';        Expression = { $lock.Connection.Host } },
                                                                                                                    @{ Name = 'InfoBase';    Expression = { @{ Descr = $lock.Connection.InfoBase.Descr; Name = $lock.Connection.InfoBase.Name } } },
                                                                                                                    @{ Name = 'SessionID';   Expression = { $lock.Connection.SessionID } },
                                                                                                                    @{ Name = 'Process';     Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object @{ Name = 'AvailablePerfomance'; Expression = { $lock.Connection.Process.AvailablePerfomance } },
                                                                                                                                                                    @{ Name = 'AvgBackCallTime';     Expression = { $lock.Connection.Process.AvgBackCallTime } },
                                                                                                                                                                    @{ Name = 'AvgCallTime';         Expression = { $lock.Connection.Process.AvgCallTime } },
                                                                                                                                                                    @{ Name = 'AvgDBCallTime';       Expression = { $lock.Connection.Process.AvgDBCallTime } },
                                                                                                                                                                    @{ Name = 'AvgLockCallTime';     Expression = { $lock.Connection.Process.AvgLockCallTime } },
                                                                                                                                                                    @{ Name = 'AvgServerCallTime';   Expression = { $lock.Connection.Process.AvgServerCallTime } },
                                                                                                                                                                    @{ Name = 'AvgThreads';          Expression = { $lock.Connection.Process.AvgThreads } },
                                                                                                                                                                    @{ Name = 'Capacity';            Expression = { $lock.Connection.Process.Capacity } },
                                                                                                                                                                    @{ Name = 'Connections';         Expression = { $lock.Connection.Process.Connections } },
                                                                                                                                                                    @{ Name = 'HostName';            Expression = { $lock.Connection.Process.HostName } },
                                                                                                                                                                    @{ Name = 'IsEnable';            Expression = { $lock.Connection.Process.IsEnable } },
                                                                                                                                                                    @{ Name = 'License';             Expression = { try { $lock.Connection.Process.License.FullPresentation } catch { $null } } },
                                                                                                                                                                    @{ Name = 'MainPort';            Expression = { $lock.Connection.Process.MainPort } },
                                                                                                                                                                    @{ Name = 'MemoryExcessTime';    Expression = { $lock.Connection.Process.MemoryExcessTime } },
                                                                                                                                                                    @{ Name = 'MemorySize';          Expression = { $lock.Connection.Process.MemorySize } },
                                                                                                                                                                    @{ Name = 'PID';                 Expression = { $lock.Connection.Process.PID } },
                                                                                                                                                                    @{ Name = 'Running';             Expression = { $lock.Connection.Process.Running } },
                                                                                                                                                                    @{ Name = 'SelectionSize';       Expression = { $lock.Connection.Process.SelectionSize } },
                                                                                                                                                                    @{ Name = 'StartedAt';           Expression = { $lock.Connection.Process.Process.StartedAt } },
                                                                                                                                                                    @{ Name = 'Use';                 Expression = { $lock.Connection.Process.Use } } } } } } } },
                                                                @{ Name = 'LockDescr'; Expression = { $lock.LockDescr } },
                                                                @{ Name = 'LockedAt';  Expression = { $lock.LockedAt } },
                                                                @{ Name = 'Object';    Expression = { $lock.Object } },
                                                                @{ Name = 'Session';   Expression = { if ( $ShowSessions -eq 'Everywhere' ) { 1| Select-Object @{ Name = 'AppID'; Expression = { $lock.Session.AppID } },
                                                                                                            @{ Name = 'blockedByDBMS';                 Expression = { $lock.Session.blockedByDBMS } },
                                                                                                            @{ Name = 'blockedByLS';                   Expression = { $lock.Session.blockedByLS } },
                                                                                                            @{ Name = 'bytesAll';                      Expression = { $lock.Session.bytesAll } },
                                                                                                            @{ Name = 'bytesLast5Min';                 Expression = { $lock.Session.bytesLast5Min } },
                                                                                                            @{ Name = 'callsAll';                      Expression = { $lock.Session.callsAll } },
                                                                                                            @{ Name = 'callsLast5Min';                 Expression = { $lock.Session.callsLast5Min } },
                                                                                                            @{ Name = 'dbmsBytesAll';                  Expression = { $lock.Session.dbmsBytesAll } },
                                                                                                            @{ Name = 'dbmsBytesLast5Min';             Expression = { $lock.Session.dbmsBytesLast5Min } },
                                                                                                            @{ Name = 'dbProcInfo';                    Expression = { $lock.Session.dbProcInfo } },
                                                                                                            @{ Name = 'dbProcTook';                    Expression = { $lock.Session.dbProcTook } },
                                                                                                            @{ Name = 'dbProcTookAt';                  Expression = { $lock.Session.dbProcTookAt } },
                                                                                                            @{ Name = 'durationAll';                   Expression = { $lock.Session.durationAll } },
                                                                                                            @{ Name = 'durationAllDBMS';               Expression = { $lock.Session.durationAllDBMS } },
                                                                                                            @{ Name = 'durationCurrent';               Expression = { $lock.Session.durationCurrent } },
                                                                                                            @{ Name = 'durationCurrentDBMS';           Expression = { $lock.Session.durationCurrentDBMS } },
                                                                                                            @{ Name = 'durationLast5Min';              Expression = { $lock.Session.durationLast5Min } },
                                                                                                            @{ Name = 'durationLast5MinDBMS';          Expression = { $lock.Session.durationLast5MinDBMS } },
                                                                                                            @{ Name = 'Hibernate';                     Expression = { $lock.Session.Hibernate } },
                                                                                                            @{ Name = 'HibernateSessionTerminateTime'; Expression = { $lock.Session.HibernateSessionTerminateTime } },
                                                                                                            @{ Name = 'Host';                          Expression = { $lock.Session.Host } },
                                                                                                            @{ Name = 'InBytesAll';                    Expression = { $lock.Session.InBytesAll } },
                                                                                                            @{ Name = 'InBytesCurrent';                Expression = { $lock.Session.InBytesCurrent } },
                                                                                                            @{ Name = 'InBytesLast5Min';               Expression = { $lock.Session.InBytesLast5Min } },
                                                                                                            @{ Name = 'InfoBase';                      Expression = { @{ Descr = $lock.Session.InfoBase.Descr; Name = $lock.Session.InfoBase.Name } } },
                                                                                                            @{ Name = 'LastActiveAt';                  Expression = { $lock.Session.LastActiveAt } },
                                                                                                            @{ Name = 'License';                       Expression = { try { $lock.Session.Process.License.FullPresentation } catch { $null } } },
                                                                                                            @{ Name = 'Locale';                        Expression = { $lock.Session.Locale } },
                                                                                                            @{ Name = 'MemoryAll';                     Expression = { $lock.Session.MemoryAll } },
                                                                                                            @{ Name = 'MemoryCurrent';                 Expression = { $lock.Session.MemoryCurrent } },
                                                                                                            @{ Name = 'MemoryLast5Min';                Expression = { $lock.Session.MemoryLast5Min } },
                                                                                                            @{ Name = 'OutBytesAll';                   Expression = { $lock.Session.OutBytesAll } },
                                                                                                            @{ Name = 'OutBytesCurrent';               Expression = { $lock.Session.OutBytesCurrent } },
                                                                                                            @{ Name = 'OutBytesLast5Min';              Expression = { $lock.Session.OutBytesLast5Min } },
                                                                                                            @{ Name = 'PassiveSessionHibernateTime';   Expression = { $lock.Session.PassiveSessionHibernateTime } },
                                                                                                            @{ Name = 'SessionID';                     Expression = { $lock.Session.SessionID } },
                                                                                                            @{ Name = 'StartedAt';                     Expression = { $lock.Session.StartedAt } },
                                                                                                            @{ Name = 'UserName';                      Expression = { $lock.Session.UserName } },
                                                                                                            @{ Name = 'Process'; Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object @{ Name = 'AvailablePerfomance'; Expression = { $lock.Session.Process.AvailablePerfomance } },
                                                                                                                                                                    @{ Name = 'AvgBackCallTime';     Expression = { $lock.Session.Process.AvgBackCallTime } },
                                                                                                                                                                    @{ Name = 'AvgCallTime';         Expression = { $lock.Session.Process.AvgCallTime } },
                                                                                                                                                                    @{ Name = 'AvgDBCallTime';       Expression = { $lock.Session.Process.AvgDBCallTime } },
                                                                                                                                                                    @{ Name = 'AvgLockCallTime';     Expression = { $lock.Session.Process.AvgLockCallTime } },
                                                                                                                                                                    @{ Name = 'AvgServerCallTime';   Expression = { $lock.Session.Process.AvgServerCallTime } },
                                                                                                                                                                    @{ Name = 'AvgThreads';          Expression = { $lock.Session.Process.AvgThreads } },
                                                                                                                                                                    @{ Name = 'Capacity';            Expression = { $lock.Session.Process.Capacity } },
                                                                                                                                                                    @{ Name = 'Connections';         Expression = { $lock.Session.Process.Connections } },
                                                                                                                                                                    @{ Name = 'HostName';            Expression = { $lock.Session.Process.HostName } },
                                                                                                                                                                    @{ Name = 'IsEnable';            Expression = { $lock.Session.Process.IsEnable } },
                                                                                                                                                                    @{ Name = 'License';             Expression = { try { $lock.Session.Process.License.FullPresentation } catch { $null } } },
                                                                                                                                                                    @{ Name = 'MainPort';            Expression = { $lock.Session.Process.MainPort   } },
                                                                                                                                                                    @{ Name = 'MemoryExcessTime';    Expression = { $lock.Session.Process.MemoryExcessTime } },
                                                                                                                                                                    @{ Name = 'MemorySize';          Expression = { $lock.Session.Process.MemorySize } },
                                                                                                                                                                    @{ Name = 'PID';                 Expression = { $lock.Session.Process.PID } },
                                                                                                                                                                    @{ Name = 'Running';             Expression = { $lock.Session.Process.Running } },
                                                                                                                                                                    @{ Name = 'SelectionSize';       Expression = { $lock.Session.Process.SelectionSize } },
                                                                                                                                                                    @{ Name = 'StartedAt';           Expression = { $lock.Session.Process.Process.StartedAt } },
                                                                                                                                                                    @{ Name = 'Use';                 Expression = { $lock.Session.Process.Use } } } } },
                                                                                                            @{ Name = 'Connection'; Expression = { if ( $ShowConnections -eq 'Everywhere' ) { 1 | Select-Object @{ Name = 'Application'; Expression = { $lock.Session.Application } },
                                                                                                                                                                        @{ Name = 'blockedByLS'; Expression = { $lock.Session.blockedByLS } },
                                                                                                                                                                        @{ Name = 'ConnectedAt'; Expression = { $lock.Session.ConnectedAt } },
                                                                                                                                                                        @{ Name = 'ConnID';      Expression = { $lock.Session.ConnID } },
                                                                                                                                                                        @{ Name = 'Host';        Expression = { $lock.Session.Host } },
                                                                                                                                                                        @{ Name = 'InfoBase';    Expression = { @{ Descr = $lock.Session.InfoBase.Descr; Name = $lock.Session.InfoBase.Name } } },
                                                                                                                                                                        @{ Name = 'SessionID';   Expression = { $lock.Session.SessionID } },
                                                                                                                                                                        @{ Name = 'Process';     Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object   @{ Name = 'AvailablePerfomance'; Expression = { $lock.Session.Process.AvailablePerfomance } },
                                                                                                                                                                                                                                    @{ Name = 'AvgBackCallTime';     Expression = { $lock.Session.Process.AvgBackCallTime } },
                                                                                                                                                                                                                                    @{ Name = 'AvgCallTime';         Expression = { $lock.Session.Process.AvgCallTime } },
                                                                                                                                                                                                                                    @{ Name = 'AvgDBCallTime';       Expression = { $lock.Session.Process.AvgDBCallTime } },
                                                                                                                                                                                                                                    @{ Name = 'AvgLockCallTime';     Expression = { $lock.Session.Process.AvgLockCallTime } },
                                                                                                                                                                                                                                    @{ Name = 'AvgServerCallTime';   Expression = { $lock.Session.Process.AvgServerCallTime } },
                                                                                                                                                                                                                                    @{ Name = 'AvgThreads';          Expression = { $lock.Session.Process.AvgThreads } },
                                                                                                                                                                                                                                    @{ Name = 'Capacity';            Expression = { $lock.Session.Process.Capacity } },
                                                                                                                                                                                                                                    @{ Name = 'Connections';         Expression = { $lock.Session.Process.Connections } },
                                                                                                                                                                                                                                    @{ Name = 'HostName';            Expression = { $lock.Session.Process.HostName } },
                                                                                                                                                                                                                                    @{ Name = 'IsEnable';            Expression = { $lock.Session.Process.IsEnable } },
                                                                                                                                                                                                                                    @{ Name = 'License';             Expression = { try { $lock.Session.Process.License.FullPresentation } catch { $null } } },
                                                                                                                                                                                                                                    @{ Name = 'MainPort';            Expression = { $lock.Session.Process.MainPort   } },
                                                                                                                                                                                                                                    @{ Name = 'MemoryExcessTime';    Expression = { $lock.Session.Process.MemoryExcessTime } },
                                                                                                                                                                                                                                    @{ Name = 'MemorySize';          Expression = { $lock.Session.Process.MemorySize } },
                                                                                                                                                                                                                                    @{ Name = 'PID';                 Expression = { $lock.Session.Process.PID } },
                                                                                                                                                                                                                                    @{ Name = 'Running';             Expression = { $lock.Session.Process.Running } },
                                                                                                                                                                                                                                    @{ Name = 'SelectionSize';       Expression = { $lock.Session.Process.SelectionSize } },
                                                                                                                                                                                                                                    @{ Name = 'StartedAt';           Expression = { $lock.Session.Process.Process.StartedAt } },
                                                                                                                                                                                                                                    @{ Name = 'Use';                 Expression = { $lock.Session.Process.Use } } } } } } } } } } }
                            }

                            Add-Member -InputObject $objCC[$objCC.Count-1] -Name ConnectionLocks -Value $objLock -MemberType NoteProperty

                        }

                    }             

                    Add-Member -InputObject $cls -Name Connections -Value $objCC -MemberType NoteProperty

                }

                if ( -not $NoInfobases ) {

                    $infoBases = $connection.GetInfoBases( $cluster )
                    $objInfoBases = @()
                    foreach( $infoBase in $infoBases ) {
    
                        $objInfoBases += 1| Select-Object @{ Name = 'Descr'; Expression =  { $infoBase.Descr } },
                                                            @{ Name = 'Name';  Expression = { $infoBase.Name } }
                            
                        if ( $ShowSessions -eq 'Everywhere' ) {

                            $infoBaseSessions = $connection.GetInfoBaseSessions( $cluster, $infoBase )
                            $objInfoBaseSession = @()
                            foreach( $ibSession in $infoBaseSessions ) {
                                $objInfoBaseSession += 1| Select-Object @{ Name = 'AppID';                         Expression = { $ibSession.AppID } },
                                                                        @{ Name = 'blockedByDBMS';                 Expression = { $ibSession.blockedByDBMS } },
                                                                        @{ Name = 'blockedByLS';                   Expression = { $ibSession.blockedByLS } },
                                                                        @{ Name = 'bytesAll';                      Expression = { $ibSession.bytesAll } },
                                                                        @{ Name = 'bytesLast5Min';                 Expression = { $ibSession.bytesLast5Min } },
                                                                        @{ Name = 'callsAll';                      Expression = { $ibSession.callsAll } },
                                                                        @{ Name = 'callsLast5Min';                 Expression = { $ibSession.callsLast5Min } },
                                                                        @{ Name = 'dbmsBytesAll';                  Expression = { $ibSession.dbmsBytesAll } },
                                                                        @{ Name = 'dbmsBytesLast5Min';             Expression = { $ibSession.dbmsBytesLast5Min } },
                                                                        @{ Name = 'dbProcInfo';                    Expression = { $ibSession.dbProcInfo } },
                                                                        @{ Name = 'dbProcTook';                    Expression = { $ibSession.dbProcTook } },
                                                                        @{ Name = 'dbProcTookAt';                  Expression = { $ibSession.dbProcTookAt } },
                                                                        @{ Name = 'durationAll';                   Expression = { $ibSession.durationAll } },
                                                                        @{ Name = 'durationAllDBMS';               Expression = { $ibSession.durationAllDBMS } },
                                                                        @{ Name = 'durationCurrent';               Expression = { $ibSession.durationCurrent } },
                                                                        @{ Name = 'durationCurrentDBMS';           Expression = { $ibSession.durationCurrentDBMS } },
                                                                        @{ Name = 'durationLast5Min';              Expression = { $ibSession.durationLast5Min } },
                                                                        @{ Name = 'durationLast5MinDBMS';          Expression = { $ibSession.durationLast5MinDBMS } },
                                                                        @{ Name = 'Hibernate';                     Expression = { $ibSession.Hibernate } },
                                                                        @{ Name = 'HibernateSessionTerminateTime'; Expression = { $ibSession.HibernateSessionTerminateTime } },
                                                                        @{ Name = 'Host';                          Expression = { $ibSession.Host } },
                                                                        @{ Name = 'InBytesAll';                    Expression = { $ibSession.InBytesAll } },
                                                                        @{ Name = 'InBytesCurrent';                Expression = { $ibSession.InBytesCurrent } },
                                                                        @{ Name = 'InBytesLast5Min';               Expression = { $ibSession.InBytesLast5Min } },
                                                                        @{ Name = 'InfoBase';                      Expression = { @{ Descr = $ibSession.InfoBase.Descr; Name = $ibSession.InfoBase.Name } } },
                                                                        @{ Name = 'LastActiveAt';                  Expression = { $ibSession.LastActiveAt } },
                                                                        @{ Name = 'License';                       Expression = { try { $ibSession.Process.License.FullPresentation } catch { $null } } },
                                                                        @{ Name = 'Locale';                        Expression = { $ibSession.Locale } },
                                                                        @{ Name = 'MemoryAll';                     Expression = { $ibSession.MemoryAll } },
                                                                        @{ Name = 'MemoryCurrent';                 Expression = { $ibSession.MemoryCurrent } },
                                                                        @{ Name = 'MemoryLast5Min';                Expression = { $ibSession.MemoryLast5Min } },
                                                                        @{ Name = 'OutBytesAll';                   Expression = { $ibSession.OutBytesAll } },
                                                                        @{ Name = 'OutBytesCurrent';               Expression = { $ibSession.OutBytesCurrent } },
                                                                        @{ Name = 'OutBytesLast5Min';              Expression = { $ibSession.OutBytesLast5Min } },
                                                                        @{ Name = 'PassiveSessionHibernateTime';   Expression = { $ibSession.PassiveSessionHibernateTime } },
                                                                        @{ Name = 'SessionID';                     Expression = { $ibSession.SessionID } },
                                                                        @{ Name = 'StartedAt';                     Expression = { $ibSession.StartedAt } },
                                                                        @{ Name = 'UserName';                      Expression = { $ibSession.UserName } },
                                                                        @{ Name = 'Process'; Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object @{ Name = 'AvailablePerfomance'; Expression = { $ibSession.Process.AvailablePerfomance } },
                                                                                                                                @{ Name = 'AvgBackCallTime';     Expression = { $ibSession.Process.AvgBackCallTime } },
                                                                                                                                @{ Name = 'AvgCallTime';         Expression = { $ibSession.Process.AvgCallTime } },
                                                                                                                                @{ Name = 'AvgDBCallTime';       Expression = { $ibSession.Process.AvgDBCallTime } },
                                                                                                                                @{ Name = 'AvgLockCallTime';     Expression = { $ibSession.Process.AvgLockCallTime } },
                                                                                                                                @{ Name = 'AvgServerCallTime';   Expression = { $ibSession.Process.AvgServerCallTime } },
                                                                                                                                @{ Name = 'AvgThreads';          Expression = { $ibSession.Process.AvgThreads } },
                                                                                                                                @{ Name = 'Capacity';            Expression = { $ibSession.Process.Capacity } },
                                                                                                                                @{ Name = 'Connections';         Expression = { $ibSession.Process.Connections } },
                                                                                                                                @{ Name = 'HostName';            Expression = { $ibSession.Process.HostName } },
                                                                                                                                @{ Name = 'IsEnable';            Expression = { $ibSession.Process.IsEnable } },
                                                                                                                                @{ Name = 'License';             Expression = { try { $ibSession.Process.License.FullPresentation } catch { $null } } },
                                                                                                                                @{ Name = 'MainPort';            Expression = { $ibSession.Process.MainPort   } },
                                                                                                                                @{ Name = 'MemoryExcessTime';    Expression = { $ibSession.Process.MemoryExcessTime } },
                                                                                                                                @{ Name = 'MemorySize';          Expression = { $ibSession.Process.MemorySize } },
                                                                                                                                @{ Name = 'PID';                 Expression = { $ibSession.Process.PID } },
                                                                                                                                @{ Name = 'Running';             Expression = { $ibSession.Process.Running } },
                                                                                                                                @{ Name = 'SelectionSize';       Expression = { $ibSession.Process.SelectionSize } },
                                                                                                                                @{ Name = 'StartedAt';           Expression = { $ibSession.Process.Process.StartedAt } },
                                                                                                                                @{ Name = 'Use';                 Expression = { $ibSession.Process.Use } } } } },
                                                                        @{ Name = 'Connection'; Expression = { if ( $ShowConnections -eq 'Everywhere' ) { 1 | Select-Object @{ Name = 'Application'; Expression = { $ibSession.Application } },
                                                                                                                                    @{ Name = 'blockedByLS'; Expression = { $ibSession.blockedByLS } },
                                                                                                                                    @{ Name = 'ConnectedAt'; Expression = { $ibSession.ConnectedAt } },
                                                                                                                                    @{ Name = 'ConnID';      Expression = { $ibSession.ConnID } },
                                                                                                                                    @{ Name = 'Host';        Expression = { $ibSession.Host } },
                                                                                                                                    @{ Name = 'InfoBase';    Expression = { @{ Descr = $ibSession.InfoBase.Descr; Name = $ibSession.InfoBase.Name } } },
                                                                                                                                    @{ Name = 'SessionID';   Expression = { $ibSession.SessionID } },
                                                                                                                                    @{ Name = 'Process';     Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object   @{ Name = 'AvailablePerfomance'; Expression = { $ibSession.Process.AvailablePerfomance } },
                                                                                                                                                                                                @{ Name = 'AvgBackCallTime';     Expression = { $ibSession.Process.AvgBackCallTime } },
                                                                                                                                                                                                @{ Name = 'AvgCallTime';         Expression = { $ibSession.Process.AvgCallTime } },
                                                                                                                                                                                                @{ Name = 'AvgDBCallTime';       Expression = { $ibSession.Process.AvgDBCallTime } },
                                                                                                                                                                                                @{ Name = 'AvgLockCallTime';     Expression = { $ibSession.Process.AvgLockCallTime } },
                                                                                                                                                                                                @{ Name = 'AvgServerCallTime';   Expression = { $ibSession.Process.AvgServerCallTime } },
                                                                                                                                                                                                @{ Name = 'AvgThreads';          Expression = { $ibSession.Process.AvgThreads } },
                                                                                                                                                                                                @{ Name = 'Capacity';            Expression = { $ibSession.Process.Capacity } },
                                                                                                                                                                                                @{ Name = 'Connections';         Expression = { $ibSession.Process.Connections } },
                                                                                                                                                                                                @{ Name = 'HostName';            Expression = { $ibSession.Process.HostName } },
                                                                                                                                                                                                @{ Name = 'IsEnable';            Expression = { $ibSession.Process.IsEnable } },
                                                                                                                                                                                                @{ Name = 'License';             Expression = { try { $ibSession.Process.License.FullPresentation } catch { $null } } },
                                                                                                                                                                                                @{ Name = 'MainPort';            Expression = { $ibSession.Process.MainPort   } },
                                                                                                                                                                                                @{ Name = 'MemoryExcessTime';    Expression = { $ibSession.Process.MemoryExcessTime } },
                                                                                                                                                                                                @{ Name = 'MemorySize';          Expression = { $ibSession.Process.MemorySize } },
                                                                                                                                                                                                @{ Name = 'PID';                 Expression = { $ibSession.Process.PID } },
                                                                                                                                                                                                @{ Name = 'Running';             Expression = { $ibSession.Process.Running } },
                                                                                                                                                                                                @{ Name = 'SelectionSize';       Expression = { $ibSession.Process.SelectionSize } },
                                                                                                                                                                                                @{ Name = 'StartedAt';           Expression = { $ibSession.Process.Process.StartedAt } },
                                                                                                                                                                                                @{ Name = 'Use';                 Expression = { $ibSession.Process.Use } } } } } } } }
                            }

                            Add-Member -InputObject $objInfoBases[$objInfoBases.Count-1] -Name InfoBaseSessions -Value $objInfoBaseSession -MemberType NoteProperty
                            
                        }

                        if ( $ShowLocks -eq 'Everywhere' ) {
                            $nfoBaseLocks = $connection.GetInfoBaseLocks( $cluster, $infoBase )
                            $objIBL = @()
                            foreach( $ibLock in $nfoBaseLocks ) {
                                $objIBL += 1 | Select-Object @{ Name = 'Connection'; Expression = { if ( $ShowConnections -eq 'Everywhere' ) { 1 | Select-Object @{ Name = 'Application'; Expression = { $ibLock.Application } },
                                                                                                                                                    @{ Name = 'blockedByLS'; Expression = { $ibLock.blockedByLS } },
                                                                                                                                                    @{ Name = 'ConnectedAt'; Expression = { $ibLock.ConnectedAt } },
                                                                                                                                                    @{ Name = 'ConnID';      Expression = { $ibLock.ConnID } },
                                                                                                                                                    @{ Name = 'Host';        Expression = { $ibLock.Host } },
                                                                                                                                                    @{ Name = 'InfoBase';    Expression = { @{ Descr = $ibLock.InfoBase.Descr; Name = $ibLock.InfoBase.Name } } },
                                                                                                                                                    @{ Name = 'SessionID';   Expression = { $ibLock.SessionID } },
                                                                                                                                                    @{ Name = 'Process';     Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object   @{ Name = 'AvailablePerfomance'; Expression = { $ibLock.Process.AvailablePerfomance } },
                                                                                                                                                                                                                @{ Name = 'AvgBackCallTime';     Expression = { $ibLock.Process.AvgBackCallTime } },
                                                                                                                                                                                                                @{ Name = 'AvgCallTime';         Expression = { $ibLock.Process.AvgCallTime } },
                                                                                                                                                                                                                @{ Name = 'AvgDBCallTime';       Expression = { $ibLock.Process.AvgDBCallTime } },
                                                                                                                                                                                                                @{ Name = 'AvgLockCallTime';     Expression = { $ibLock.Process.AvgLockCallTime } },
                                                                                                                                                                                                                @{ Name = 'AvgServerCallTime';   Expression = { $ibLock.Process.AvgServerCallTime } },
                                                                                                                                                                                                                @{ Name = 'AvgThreads';          Expression = { $ibLock.Process.AvgThreads } },
                                                                                                                                                                                                                @{ Name = 'Capacity';            Expression = { $ibLock.Process.Capacity } },
                                                                                                                                                                                                                @{ Name = 'Connections';         Expression = { $ibLock.Process.Connections } },
                                                                                                                                                                                                                @{ Name = 'HostName';            Expression = { $ibLock.Process.HostName } },
                                                                                                                                                                                                                @{ Name = 'IsEnable';            Expression = { $ibLock.Process.IsEnable } },
                                                                                                                                                                                                                @{ Name = 'License';             Expression = { try { $ibLock.Process.License.FullPresentation } catch { $null } } },
                                                                                                                                                                                                                @{ Name = 'MainPort';            Expression = { $ibLock.Process.MainPort   } },
                                                                                                                                                                                                                @{ Name = 'MemoryExcessTime';    Expression = { $ibLock.Process.MemoryExcessTime } },
                                                                                                                                                                                                                @{ Name = 'MemorySize';          Expression = { $ibLock.Process.MemorySize } },
                                                                                                                                                                                                                @{ Name = 'PID';                 Expression = { $ibLock.Process.PID } },
                                                                                                                                                                                                                @{ Name = 'Running';             Expression = { $ibLock.Process.Running } },
                                                                                                                                                                                                                @{ Name = 'SelectionSize';       Expression = { $ibLock.Process.SelectionSize } },
                                                                                                                                                                                                                @{ Name = 'StartedAt';           Expression = { $ibLock.Process.Process.StartedAt } },
                                                                                                                                                                                                                @{ Name = 'Use';                 Expression = { $ibLock.Process.Use } } } } } } } },
                                                            @{ Name = 'LockDescr'; Expression = { $ibLock.LockDescr } },
                                                            @{ Name = 'LockedAt';  Expression = { $ibLock.MainManager } },
                                                            @{ Name = 'Object';    Expression = { $ibLock.MainPort } },
                                                            @{ Name = 'Session';   Expression = { if ( $ShowSessions -eq 'Everywhere' ) { 1| Select-Object @{ Name = 'AppID'; Expression = { $ibLock.Session.AppID } },
                                                                                                            @{ Name = 'blockedByDBMS';                 Expression = { $ibLock.Session.blockedByDBMS } },
                                                                                                            @{ Name = 'blockedByLS';                   Expression = { $ibLock.Session.blockedByLS } },
                                                                                                            @{ Name = 'bytesAll';                      Expression = { $ibLock.Session.bytesAll } },
                                                                                                            @{ Name = 'bytesLast5Min';                 Expression = { $ibLock.Session.bytesLast5Min } },
                                                                                                            @{ Name = 'callsAll';                      Expression = { $ibLock.Session.callsAll } },
                                                                                                            @{ Name = 'callsLast5Min';                 Expression = { $ibLock.Session.callsLast5Min } },
                                                                                                            @{ Name = 'dbmsBytesAll';                  Expression = { $ibLock.Session.dbmsBytesAll } },
                                                                                                            @{ Name = 'dbmsBytesLast5Min';             Expression = { $ibLock.Session.dbmsBytesLast5Min } },
                                                                                                            @{ Name = 'dbProcInfo';                    Expression = { $ibLock.Session.dbProcInfo } },
                                                                                                            @{ Name = 'dbProcTook';                    Expression = { $ibLock.Session.dbProcTook } },
                                                                                                            @{ Name = 'dbProcTookAt';                  Expression = { $ibLock.Session.dbProcTookAt } },
                                                                                                            @{ Name = 'durationAll';                   Expression = { $ibLock.Session.durationAll } },
                                                                                                            @{ Name = 'durationAllDBMS';               Expression = { $ibLock.Session.durationAllDBMS } },
                                                                                                            @{ Name = 'durationCurrent';               Expression = { $ibLock.Session.durationCurrent } },
                                                                                                            @{ Name = 'durationCurrentDBMS';           Expression = { $ibLock.Session.durationCurrentDBMS } },
                                                                                                            @{ Name = 'durationLast5Min';              Expression = { $ibLock.Session.durationLast5Min } },
                                                                                                            @{ Name = 'durationLast5MinDBMS';          Expression = { $ibLock.Session.durationLast5MinDBMS } },
                                                                                                            @{ Name = 'Hibernate';                     Expression = { $ibLock.Session.Hibernate } },
                                                                                                            @{ Name = 'HibernateSessionTerminateTime'; Expression = { $ibLock.Session.HibernateSessionTerminateTime } },
                                                                                                            @{ Name = 'Host';                          Expression = { $ibLock.Session.Host } },
                                                                                                            @{ Name = 'InBytesAll';                    Expression = { $ibLock.Session.InBytesAll } },
                                                                                                            @{ Name = 'InBytesCurrent';                Expression = { $ibLock.Session.InBytesCurrent } },
                                                                                                            @{ Name = 'InBytesLast5Min';               Expression = { $ibLock.Session.InBytesLast5Min } },
                                                                                                            @{ Name = 'InfoBase';                      Expression = { @{ Descr = $ibLock.Session.InfoBase.Descr; Name = $ibLock.Session.InfoBase.Name } } },
                                                                                                            @{ Name = 'LastActiveAt';                  Expression = { $ibLock.Session.LastActiveAt } },
                                                                                                            @{ Name = 'License';                       Expression = { try { $ibLock.Session.Process.License.FullPresentation } catch { $null } } },
                                                                                                            @{ Name = 'Locale';                        Expression = { $ibLock.Session.Locale } },
                                                                                                            @{ Name = 'MemoryAll';                     Expression = { $ibLock.Session.MemoryAll } },
                                                                                                            @{ Name = 'MemoryCurrent';                 Expression = { $ibLock.Session.MemoryCurrent } },
                                                                                                            @{ Name = 'MemoryLast5Min';                Expression = { $ibLock.Session.MemoryLast5Min } },
                                                                                                            @{ Name = 'OutBytesAll';                   Expression = { $ibLock.Session.OutBytesAll } },
                                                                                                            @{ Name = 'OutBytesCurrent';               Expression = { $ibLock.Session.OutBytesCurrent } },
                                                                                                            @{ Name = 'OutBytesLast5Min';              Expression = { $ibLock.Session.OutBytesLast5Min } },
                                                                                                            @{ Name = 'PassiveSessionHibernateTime';   Expression = { $ibLock.Session.PassiveSessionHibernateTime } },
                                                                                                            @{ Name = 'SessionID';                     Expression = { $ibLock.Session.SessionID } },
                                                                                                            @{ Name = 'StartedAt';                     Expression = { $ibLock.Session.StartedAt } },
                                                                                                            @{ Name = 'UserName';                      Expression = { $ibLock.Session.UserName } },
                                                                                                            @{ Name = 'Process'; Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object @{ Name = 'AvailablePerfomance'; Expression = { $ibLock.Session.Process.AvailablePerfomance } },
                                                                                                                                                                    @{ Name = 'AvgBackCallTime';     Expression = { $ibLock.Session.Process.AvgBackCallTime } },
                                                                                                                                                                    @{ Name = 'AvgCallTime';         Expression = { $ibLock.Session.Process.AvgCallTime } },
                                                                                                                                                                    @{ Name = 'AvgDBCallTime';       Expression = { $ibLock.Session.Process.AvgDBCallTime } },
                                                                                                                                                                    @{ Name = 'AvgLockCallTime';     Expression = { $ibLock.Session.Process.AvgLockCallTime } },
                                                                                                                                                                    @{ Name = 'AvgServerCallTime';   Expression = { $ibLock.Session.Process.AvgServerCallTime } },
                                                                                                                                                                    @{ Name = 'AvgThreads';          Expression = { $ibLock.Session.Process.AvgThreads } },
                                                                                                                                                                    @{ Name = 'Capacity';            Expression = { $ibLock.Session.Process.Capacity } },
                                                                                                                                                                    @{ Name = 'Connections';         Expression = { $ibLock.Session.Process.Connections } },
                                                                                                                                                                    @{ Name = 'HostName';            Expression = { $ibLock.Session.Process.HostName } },
                                                                                                                                                                    @{ Name = 'IsEnable';            Expression = { $ibLock.Session.Process.IsEnable } },
                                                                                                                                                                    @{ Name = 'License';             Expression = { try { $ibLock.Session.Process.License.FullPresentation } catch { $null } } },
                                                                                                                                                                    @{ Name = 'MainPort';            Expression = { $ibLock.Session.Process.MainPort   } },
                                                                                                                                                                    @{ Name = 'MemoryExcessTime';    Expression = { $ibLock.Session.Process.MemoryExcessTime } },
                                                                                                                                                                    @{ Name = 'MemorySize';          Expression = { $ibLock.Session.Process.MemorySize } },
                                                                                                                                                                    @{ Name = 'PID';                 Expression = { $ibLock.Session.Process.PID } },
                                                                                                                                                                    @{ Name = 'Running';             Expression = { $ibLock.Session.Process.Running } },
                                                                                                                                                                    @{ Name = 'SelectionSize';       Expression = { $ibLock.Session.Process.SelectionSize } },
                                                                                                                                                                    @{ Name = 'StartedAt';           Expression = { $ibLock.Session.Process.Process.StartedAt } },
                                                                                                                                                                    @{ Name = 'Use';                 Expression = { $ibLock.Session.Process.Use } } } } },
                                                                                                            @{ Name = 'Connection'; Expression = { if ( $ShowConnections -eq 'Everywhere') { 1 | Select-Object @{ Name = 'Application'; Expression = { $ibLock.Session.Application } },
                                                                                                                                                                        @{ Name = 'blockedByLS'; Expression = { $ibLock.Session.blockedByLS } },
                                                                                                                                                                        @{ Name = 'ConnectedAt'; Expression = { $ibLock.Session.ConnectedAt } },
                                                                                                                                                                        @{ Name = 'ConnID';      Expression = { $ibLock.Session.ConnID } },
                                                                                                                                                                        @{ Name = 'Host';        Expression = { $ibLock.Session.Host } },
                                                                                                                                                                        @{ Name = 'InfoBase';    Expression = { @{ Descr = $ibLock.Session.InfoBase.Descr; Name = $ibLock.Session.InfoBase.Name } } },
                                                                                                                                                                        @{ Name = 'SessionID';   Expression = { $ibLock.Session.SessionID } },
                                                                                                                                                                        @{ Name = 'Process';     Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object   @{ Name = 'AvailablePerfomance'; Expression = { $ibLock.Session.Process.AvailablePerfomance } },
                                                                                                                                                                                                                                    @{ Name = 'AvgBackCallTime';     Expression = { $ibLock.Session.Process.AvgBackCallTime } },
                                                                                                                                                                                                                                    @{ Name = 'AvgCallTime';         Expression = { $ibLock.Session.Process.AvgCallTime } },
                                                                                                                                                                                                                                    @{ Name = 'AvgDBCallTime';       Expression = { $ibLock.Session.Process.AvgDBCallTime } },
                                                                                                                                                                                                                                    @{ Name = 'AvgLockCallTime';     Expression = { $ibLock.Session.Process.AvgLockCallTime } },
                                                                                                                                                                                                                                    @{ Name = 'AvgServerCallTime';   Expression = { $ibLock.Session.Process.AvgServerCallTime } },
                                                                                                                                                                                                                                    @{ Name = 'AvgThreads';          Expression = { $ibLock.Session.Process.AvgThreads } },
                                                                                                                                                                                                                                    @{ Name = 'Capacity';            Expression = { $ibLock.Session.Process.Capacity } },
                                                                                                                                                                                                                                    @{ Name = 'Connections';         Expression = { $ibLock.Session.Process.Connections } },
                                                                                                                                                                                                                                    @{ Name = 'HostName';            Expression = { $ibLock.Session.Process.HostName } },
                                                                                                                                                                                                                                    @{ Name = 'IsEnable';            Expression = { $ibLock.Session.Process.IsEnable } },
                                                                                                                                                                                                                                    @{ Name = 'License';             Expression = { try { $ibLock.Session.Process.License.FullPresentation } catch { $null } } },
                                                                                                                                                                                                                                    @{ Name = 'MainPort';            Expression = { $ibLock.Session.Process.MainPort   } },
                                                                                                                                                                                                                                    @{ Name = 'MemoryExcessTime';    Expression = { $ibLock.Session.Process.MemoryExcessTime } },
                                                                                                                                                                                                                                    @{ Name = 'MemorySize';          Expression = { $ibLock.Session.Process.MemorySize } },
                                                                                                                                                                                                                                    @{ Name = 'PID';                 Expression = { $ibLock.Session.Process.PID } },
                                                                                                                                                                                                                                    @{ Name = 'Running';             Expression = { $ibLock.Session.Process.Running } },
                                                                                                                                                                                                                                    @{ Name = 'SelectionSize';       Expression = { $ibLock.Session.Process.SelectionSize } },
                                                                                                                                                                                                                                    @{ Name = 'StartedAt';           Expression = { $ibLock.Session.Process.Process.StartedAt } },
                                                                                                                                                                                                                                    @{ Name = 'Use';                 Expression = { $ibLock.Session.Process.Use } } } } } } } } } } }
                            }
                            Add-Member -InputObject $objInfoBases[$objInfoBases.Count-1] -Name InfoBaseLocks -Value $objIBL -MemberType NoteProperty
                        }

                        if ( $ShowConnections -eq 'Everywhere' ) {
                            $nfoBaseConnections = $connection.GetInfoBaseConnections( $cluster, $infoBase )
                            $objIBC = @()
                            foreach( $ibConnection in $nfoBaseConnections ) {
                                $objIBC += 1 | Select-Object @{ Name = 'Application'; Expression = { $ibConnection.Application } },
                                                            @{ Name = 'blockedByLS'; Expression = { $ibConnection.blockedByLS } },
                                                            @{ Name = 'ConnectedAt'; Expression = { $ibConnection.ConnectedAt } },
                                                            @{ Name = 'ConnID';      Expression = { $ibConnection.ConnID } },
                                                            @{ Name = 'Host';        Expression = { $ibConnection.Host } },
                                                            @{ Name = 'InfoBase';    Expression = { @{ Descr = $ibConnection.InfoBase.Descr; Name = $ibConnection.InfoBase.Name } } },
                                                            @{ Name = 'SessionID';   Expression = { $ibConnection.SessionID } },
                                                            @{ Name = 'Process';     Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object   @{ Name = 'AvailablePerfomance'; Expression = { $ibConnection.Process.AvailablePerfomance } },
                                                                                                                        @{ Name = 'AvgBackCallTime';     Expression = { $ibConnection.Process.AvgBackCallTime } },
                                                                                                                        @{ Name = 'AvgCallTime';         Expression = { $ibConnection.Process.AvgCallTime } },
                                                                                                                        @{ Name = 'AvgDBCallTime';       Expression = { $ibConnection.Process.AvgDBCallTime } },
                                                                                                                        @{ Name = 'AvgLockCallTime';     Expression = { $ibConnection.Process.AvgLockCallTime } },
                                                                                                                        @{ Name = 'AvgServerCallTime';   Expression = { $ibConnection.Process.AvgServerCallTime } },
                                                                                                                        @{ Name = 'AvgThreads';          Expression = { $ibConnection.Process.AvgThreads } },
                                                                                                                        @{ Name = 'Capacity';            Expression = { $ibConnection.Process.Capacity } },
                                                                                                                        @{ Name = 'Connections';         Expression = { $ibConnection.Process.Connections } },
                                                                                                                        @{ Name = 'HostName';            Expression = { $ibConnection.Process.HostName } },
                                                                                                                        @{ Name = 'IsEnable';            Expression = { $ibConnection.Process.IsEnable } },
                                                                                                                        @{ Name = 'License';             Expression = { try { $ibConnection.Process.License.FullPresentation } catch { $null } } },
                                                                                                                        @{ Name = 'MainPort';            Expression = { $ibConnection.Process.MainPort   } },
                                                                                                                        @{ Name = 'MemoryExcessTime';    Expression = { $ibConnection.Process.MemoryExcessTime } },
                                                                                                                        @{ Name = 'MemorySize';          Expression = { $ibConnection.Process.MemorySize } },
                                                                                                                        @{ Name = 'PID';                 Expression = { $ibConnection.Process.PID } },
                                                                                                                        @{ Name = 'Running';             Expression = { $ibConnection.Process.Running } },
                                                                                                                        @{ Name = 'SelectionSize';       Expression = { $ibConnection.Process.SelectionSize } },
                                                                                                                        @{ Name = 'StartedAt';           Expression = { $ibConnection.Process.Process.StartedAt } },
                                                                                                                        @{ Name = 'Use';                 Expression = { $ibConnection.Process.Use } } } } }
                            }
                            Add-Member -InputObject $objInfoBases[$objInfoBases.Count-1] -Name InfoBaseConnections -Value $objIBC -MemberType NoteProperty
                        }

                    }

                    Add-Member -InputObject $cls -Name InfoBases -Value $objInfoBases -MemberType NoteProperty

                }

                if ( $ShowLocks -ne 'None' ) {

                    $clusterLocks = $connection.GetLocks( $cluster )
                    $objClLock = @()
                    foreach( $clLock in $clusterLocks ) {
                        $objClLock += 1 | Select-Object @{ Name = 'Connection';  Expression = { if ( $clLock.Connection ) { 1 | Select-Object @{ Name = 'Application'; Expression = { $clLock.Connection.Application } },
                                                                                                            @{ Name = 'blockedByLS'; Expression = { $clLock.Connection.blockedByLS } },
                                                                                                            @{ Name = 'ConnectedAt'; Expression = { $clLock.Connection.ConnectedAt } },
                                                                                                            @{ Name = 'ConnID';      Expression = { $clLock.Connection.ConnID } },
                                                                                                            @{ Name = 'Host';        Expression = { $clLock.Connection.Host } },
                                                                                                            @{ Name = 'InfoBase';    Expression = { @{ Descr = $clLock.Connection.InfoBase.Descr; Name = $clLock.Connection.InfoBase.Name } } },
                                                                                                            @{ Name = 'SessionID';   Expression = { $clLock.Connection.SessionID } },
                                                                                                            @{ Name = 'Process';     Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object @{ Name = 'AvailablePerfomance'; Expression = { $clLock.Connection.Process.AvailablePerfomance } },
                                                                                                                                                            @{ Name = 'AvgBackCallTime';     Expression = { $clLock.Connection.Process.AvgBackCallTime } },
                                                                                                                                                            @{ Name = 'AvgCallTime';         Expression = { $clLock.Connection.Process.AvgCallTime } },
                                                                                                                                                            @{ Name = 'AvgDBCallTime';       Expression = { $clLock.Connection.Process.AvgDBCallTime } },
                                                                                                                                                            @{ Name = 'AvgLockCallTime';     Expression = { $clLock.Connection.Process.AvgLockCallTime } },
                                                                                                                                                            @{ Name = 'AvgServerCallTime';   Expression = { $clLock.Connection.Process.AvgServerCallTime } },
                                                                                                                                                            @{ Name = 'AvgThreads';          Expression = { $clLock.Connection.Process.AvgThreads } },
                                                                                                                                                            @{ Name = 'Capacity';            Expression = { $clLock.Connection.Process.Capacity } },
                                                                                                                                                            @{ Name = 'Connections';         Expression = { $clLock.Connection.Process.Connections } },
                                                                                                                                                            @{ Name = 'HostName';            Expression = { $clLock.Connection.Process.HostName } },
                                                                                                                                                            @{ Name = 'IsEnable';            Expression = { $clLock.Connection.Process.IsEnable } },
                                                                                                                                                            @{ Name = 'License';             Expression = { try { $clLock.Connection.Process.License.FullPresentation } catch { $null } } },
                                                                                                                                                            @{ Name = 'MainPort';            Expression = { $clLock.Connection.Process.MainPort } },
                                                                                                                                                            @{ Name = 'MemoryExcessTime';    Expression = { $clLock.Connection.Process.MemoryExcessTime } },
                                                                                                                                                            @{ Name = 'MemorySize';          Expression = { $clLock.Connection.Process.MemorySize } },
                                                                                                                                                            @{ Name = 'PID';                 Expression = { $clLock.Connection.Process.PID } },
                                                                                                                                                            @{ Name = 'Running';             Expression = { $clLock.Connection.Process.Running } },
                                                                                                                                                            @{ Name = 'SelectionSize';       Expression = { $clLock.Connection.Process.SelectionSize } },
                                                                                                                                                            @{ Name = 'StartedAt';           Expression = { $clLock.Connection.Process.Process.StartedAt } },
                                                                                                                                                            @{ Name = 'Use';                 Expression = { $clLock.Connection.Process.Use } } } } } } } },
                                                        @{ Name = 'LockDescr'; Expression = { $clLock.LockDescr } },
                                                        @{ Name = 'LockedAt';  Expression = { $clLock.LockedAt } },
                                                        @{ Name = 'Object';    Expression = { $clLock.Object } },
                                                        @{ Name = 'Session';   Expression = { if ( $ShowSessions -eq 'Everywhere' ) { 1| Select-Object @{ Name = 'AppID'; Expression = { $clLock.Session.AppID } },
                                                                                                    @{ Name = 'blockedByDBMS';                 Expression = { $clLock.Session.blockedByDBMS } },
                                                                                                    @{ Name = 'blockedByLS';                   Expression = { $clLock.Session.blockedByLS } },
                                                                                                    @{ Name = 'bytesAll';                      Expression = { $clLock.Session.bytesAll } },
                                                                                                    @{ Name = 'bytesLast5Min';                 Expression = { $clLock.Session.bytesLast5Min } },
                                                                                                    @{ Name = 'callsAll';                      Expression = { $clLock.Session.callsAll } },
                                                                                                    @{ Name = 'callsLast5Min';                 Expression = { $clLock.Session.callsLast5Min } },
                                                                                                    @{ Name = 'dbmsBytesAll';                  Expression = { $clLock.Session.dbmsBytesAll } },
                                                                                                    @{ Name = 'dbmsBytesLast5Min';             Expression = { $clLock.Session.dbmsBytesLast5Min } },
                                                                                                    @{ Name = 'dbProcInfo';                    Expression = { $clLock.Session.dbProcInfo } },
                                                                                                    @{ Name = 'dbProcTook';                    Expression = { $clLock.Session.dbProcTook } },
                                                                                                    @{ Name = 'dbProcTookAt';                  Expression = { $clLock.Session.dbProcTookAt } },
                                                                                                    @{ Name = 'durationAll';                   Expression = { $clLock.Session.durationAll } },
                                                                                                    @{ Name = 'durationAllDBMS';               Expression = { $clLock.Session.durationAllDBMS } },
                                                                                                    @{ Name = 'durationCurrent';               Expression = { $clLock.Session.durationCurrent } },
                                                                                                    @{ Name = 'durationCurrentDBMS';           Expression = { $clLock.Session.durationCurrentDBMS } },
                                                                                                    @{ Name = 'durationLast5Min';              Expression = { $clLock.Session.durationLast5Min } },
                                                                                                    @{ Name = 'durationLast5MinDBMS';          Expression = { $clLock.Session.durationLast5MinDBMS } },
                                                                                                    @{ Name = 'Hibernate';                     Expression = { $clLock.Session.Hibernate } },
                                                                                                    @{ Name = 'HibernateSessionTerminateTime'; Expression = { $clLock.Session.HibernateSessionTerminateTime } },
                                                                                                    @{ Name = 'Host';                          Expression = { $clLock.Session.Host } },
                                                                                                    @{ Name = 'InBytesAll';                    Expression = { $clLock.Session.InBytesAll } },
                                                                                                    @{ Name = 'InBytesCurrent';                Expression = { $clLock.Session.InBytesCurrent } },
                                                                                                    @{ Name = 'InBytesLast5Min';               Expression = { $clLock.Session.InBytesLast5Min } },
                                                                                                    @{ Name = 'InfoBase';                      Expression = { @{ Descr = $clLock.Session.InfoBase.Descr; Name = $clLock.Session.InfoBase.Name } } },
                                                                                                    @{ Name = 'LastActiveAt';                  Expression = { $clLock.Session.LastActiveAt } },
                                                                                                    @{ Name = 'License';                       Expression = { try { $clLock.Session.Process.License.FullPresentation } catch { $null } } },
                                                                                                    @{ Name = 'Locale';                        Expression = { $clLock.Session.Locale } },
                                                                                                    @{ Name = 'MemoryAll';                     Expression = { $clLock.Session.MemoryAll } },
                                                                                                    @{ Name = 'MemoryCurrent';                 Expression = { $clLock.Session.MemoryCurrent } },
                                                                                                    @{ Name = 'MemoryLast5Min';                Expression = { $clLock.Session.MemoryLast5Min } },
                                                                                                    @{ Name = 'OutBytesAll';                   Expression = { $clLock.Session.OutBytesAll } },
                                                                                                    @{ Name = 'OutBytesCurrent';               Expression = { $clLock.Session.OutBytesCurrent } },
                                                                                                    @{ Name = 'OutBytesLast5Min';              Expression = { $clLock.Session.OutBytesLast5Min } },
                                                                                                    @{ Name = 'PassiveSessionHibernateTime';   Expression = { $clLock.Session.PassiveSessionHibernateTime } },
                                                                                                    @{ Name = 'SessionID';                     Expression = { $clLock.Session.SessionID } },
                                                                                                    @{ Name = 'StartedAt';                     Expression = { $clLock.Session.StartedAt } },
                                                                                                    @{ Name = 'UserName';                      Expression = { $clLock.Session.UserName } },
                                                                                                    @{ Name = 'Process'; Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object @{ Name = 'AvailablePerfomance'; Expression = { $clLock.Session.Process.AvailablePerfomance } },
                                                                                                                                                            @{ Name = 'AvgBackCallTime';     Expression = { $clLock.Session.Process.AvgBackCallTime } },
                                                                                                                                                            @{ Name = 'AvgCallTime';         Expression = { $clLock.Session.Process.AvgCallTime } },
                                                                                                                                                            @{ Name = 'AvgDBCallTime';       Expression = { $clLock.Session.Process.AvgDBCallTime } },
                                                                                                                                                            @{ Name = 'AvgLockCallTime';     Expression = { $clLock.Session.Process.AvgLockCallTime } },
                                                                                                                                                            @{ Name = 'AvgServerCallTime';   Expression = { $clLock.Session.Process.AvgServerCallTime } },
                                                                                                                                                            @{ Name = 'AvgThreads';          Expression = { $clLock.Session.Process.AvgThreads } },
                                                                                                                                                            @{ Name = 'Capacity';            Expression = { $clLock.Session.Process.Capacity } },
                                                                                                                                                            @{ Name = 'Connections';         Expression = { $clLock.Session.Process.Connections } },
                                                                                                                                                            @{ Name = 'HostName';            Expression = { $clLock.Session.Process.HostName } },
                                                                                                                                                            @{ Name = 'IsEnable';            Expression = { $clLock.Session.Process.IsEnable } },
                                                                                                                                                            @{ Name = 'License';             Expression = { try { $clLock.Session.Process.License.FullPresentation } catch { $null } } },
                                                                                                                                                            @{ Name = 'MainPort';            Expression = { $clLock.Session.Process.MainPort   } },
                                                                                                                                                            @{ Name = 'MemoryExcessTime';    Expression = { $clLock.Session.Process.MemoryExcessTime } },
                                                                                                                                                            @{ Name = 'MemorySize';          Expression = { $clLock.Session.Process.MemorySize } },
                                                                                                                                                            @{ Name = 'PID';                 Expression = { $clLock.Session.Process.PID } },
                                                                                                                                                            @{ Name = 'Running';             Expression = { $clLock.Session.Process.Running } },
                                                                                                                                                            @{ Name = 'SelectionSize';       Expression = { $clLock.Session.Process.SelectionSize } },
                                                                                                                                                            @{ Name = 'StartedAt';           Expression = { $clLock.Session.Process.Process.StartedAt } },
                                                                                                                                                            @{ Name = 'Use';                 Expression = { $clLock.Session.Process.Use } } } } },
                                                                                                    @{ Name = 'Connection'; Expression = { if ( $ShowConnections -eq 'Everywhere') { 1 | Select-Object @{ Name = 'Application'; Expression = { $clLock.Session.Application } },
                                                                                                                                                                @{ Name = 'blockedByLS'; Expression = { $clLock.Session.blockedByLS } },
                                                                                                                                                                @{ Name = 'ConnectedAt'; Expression = { $clLock.Session.ConnectedAt } },
                                                                                                                                                                @{ Name = 'ConnID';      Expression = { $clLock.Session.ConnID } },
                                                                                                                                                                @{ Name = 'Host';        Expression = { $clLock.Session.Host } },
                                                                                                                                                                @{ Name = 'InfoBase';    Expression = { @{ Descr = $clLock.Session.InfoBase.Descr; Name = $clLock.Session.InfoBase.Name } } },
                                                                                                                                                                @{ Name = 'SessionID';   Expression = { $clLock.Session.SessionID } },
                                                                                                                                                                @{ Name = 'Process';     Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object   @{ Name = 'AvailablePerfomance'; Expression = { $clLock.Session.Process.AvailablePerfomance } },
                                                                                                                                                                                                                            @{ Name = 'AvgBackCallTime';     Expression = { $clLock.Session.Process.AvgBackCallTime } },
                                                                                                                                                                                                                            @{ Name = 'AvgCallTime';         Expression = { $clLock.Session.Process.AvgCallTime } },
                                                                                                                                                                                                                            @{ Name = 'AvgDBCallTime';       Expression = { $clLock.Session.Process.AvgDBCallTime } },
                                                                                                                                                                                                                            @{ Name = 'AvgLockCallTime';     Expression = { $clLock.Session.Process.AvgLockCallTime } },
                                                                                                                                                                                                                            @{ Name = 'AvgServerCallTime';   Expression = { $clLock.Session.Process.AvgServerCallTime } },
                                                                                                                                                                                                                            @{ Name = 'AvgThreads';          Expression = { $clLock.Session.Process.AvgThreads } },
                                                                                                                                                                                                                            @{ Name = 'Capacity';            Expression = { $clLock.Session.Process.Capacity } },
                                                                                                                                                                                                                            @{ Name = 'Connections';         Expression = { $clLock.Session.Process.Connections } },
                                                                                                                                                                                                                            @{ Name = 'HostName';            Expression = { $clLock.Session.Process.HostName } },
                                                                                                                                                                                                                            @{ Name = 'IsEnable';            Expression = { $clLock.Session.Process.IsEnable } },
                                                                                                                                                                                                                            @{ Name = 'License';             Expression = { try { $clLock.Session.Process.License.FullPresentation } catch { $null } } },
                                                                                                                                                                                                                            @{ Name = 'MainPort';            Expression = { $clLock.Session.Process.MainPort   } },
                                                                                                                                                                                                                            @{ Name = 'MemoryExcessTime';    Expression = { $clLock.Session.Process.MemoryExcessTime } },
                                                                                                                                                                                                                            @{ Name = 'MemorySize';          Expression = { $clLock.Session.Process.MemorySize } },
                                                                                                                                                                                                                            @{ Name = 'PID';                 Expression = { $clLock.Session.Process.PID } },
                                                                                                                                                                                                                            @{ Name = 'Running';             Expression = { $clLock.Session.Process.Running } },
                                                                                                                                                                                                                            @{ Name = 'SelectionSize';       Expression = { $clLock.Session.Process.SelectionSize } },
                                                                                                                                                                                                                            @{ Name = 'StartedAt';           Expression = { $clLock.Session.Process.Process.StartedAt } },
                                                                                                                                                                                                                            @{ Name = 'Use';                 Expression = { $clLock.Session.Process.Use } } } } } } } } } } }
                    }

                    Add-Member -InputObject $cls -Name Locks -Value $objClLock -MemberType NoteProperty

                }

                if ( $ShowSessions -ne 'None' ) {
                        
                    $clusterSessions = $connection.GetSessions( $cluster )
                    $objClSession = @()
                    foreach ( $clusterSession in $clusterSessions ) {
                        $objClSession += 1| Select-Object @{ Name = 'AppID'; Expression = { $clusterSession.AppID } },
                                                        @{ Name = 'blockedByDBMS';                 Expression = { $clusterSession.blockedByDBMS } },
                                                        @{ Name = 'blockedByLS';                   Expression = { $clusterSession.blockedByLS } },
                                                        @{ Name = 'bytesAll';                      Expression = { $clusterSession.bytesAll } },
                                                        @{ Name = 'bytesLast5Min';                 Expression = { $clusterSession.bytesLast5Min } },
                                                        @{ Name = 'callsAll';                      Expression = { $clusterSession.callsAll } },
                                                        @{ Name = 'callsLast5Min';                 Expression = { $clusterSession.callsLast5Min } },
                                                        @{ Name = 'dbmsBytesAll';                  Expression = { $clusterSession.dbmsBytesAll } },
                                                        @{ Name = 'dbmsBytesLast5Min';             Expression = { $clusterSession.dbmsBytesLast5Min } },
                                                        @{ Name = 'dbProcInfo';                    Expression = { $clusterSession.dbProcInfo } },
                                                        @{ Name = 'dbProcTook';                    Expression = { $clusterSession.dbProcTook } },
                                                        @{ Name = 'dbProcTookAt';                  Expression = { $clusterSession.dbProcTookAt } },
                                                        @{ Name = 'durationAll';                   Expression = { $clusterSession.durationAll } },
                                                        @{ Name = 'durationAllDBMS';               Expression = { $clusterSession.durationAllDBMS } },
                                                        @{ Name = 'durationCurrent';               Expression = { $clusterSession.durationCurrent } },
                                                        @{ Name = 'durationCurrentDBMS';           Expression = { $clusterSession.durationCurrentDBMS } },
                                                        @{ Name = 'durationLast5Min';              Expression = { $clusterSession.durationLast5Min } },
                                                        @{ Name = 'durationLast5MinDBMS';          Expression = { $clusterSession.durationLast5MinDBMS } },
                                                        @{ Name = 'Hibernate';                     Expression = { $clusterSession.Hibernate } },
                                                        @{ Name = 'HibernateSessionTerminateTime'; Expression = { $clusterSession.HibernateSessionTerminateTime } },
                                                        @{ Name = 'Host';                          Expression = { $clusterSession.Host } },
                                                        @{ Name = 'InBytesAll';                    Expression = { $clusterSession.InBytesAll } },
                                                        @{ Name = 'InBytesCurrent';                Expression = { $clusterSession.InBytesCurrent } },
                                                        @{ Name = 'InBytesLast5Min';               Expression = { $clusterSession.InBytesLast5Min } },
                                                        @{ Name = 'InfoBase';                      Expression = { @{ Descr = $clusterSession.InfoBase.Descr; Name = $clusterSession.InfoBase.Name } } },
                                                        @{ Name = 'LastActiveAt';                  Expression = { $clusterSession.LastActiveAt } },
                                                        @{ Name = 'License';                       Expression = { try { $clusterSession.Process.License.FullPresentation } catch { $null } } },
                                                        @{ Name = 'Locale';                        Expression = { $clusterSession.Locale } },
                                                        @{ Name = 'MemoryAll';                     Expression = { $clusterSession.MemoryAll } },
                                                        @{ Name = 'MemoryCurrent';                 Expression = { $clusterSession.MemoryCurrent } },
                                                        @{ Name = 'MemoryLast5Min';                Expression = { $clusterSession.MemoryLast5Min } },
                                                        @{ Name = 'OutBytesAll';                   Expression = { $clusterSession.OutBytesAll } },
                                                        @{ Name = 'OutBytesCurrent';               Expression = { $clusterSession.OutBytesCurrent } },
                                                        @{ Name = 'OutBytesLast5Min';              Expression = { $clusterSession.OutBytesLast5Min } },
                                                        @{ Name = 'PassiveSessionHibernateTime';   Expression = { $clusterSession.PassiveSessionHibernateTime } },
                                                        @{ Name = 'SessionID';                     Expression = { $clusterSession.SessionID } },
                                                        @{ Name = 'StartedAt';                     Expression = { $clusterSession.StartedAt } },
                                                        @{ Name = 'UserName';                      Expression = { $clusterSession.UserName } },
                                                        @{ Name = 'Process'; Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object @{ Name = 'AvailablePerfomance'; Expression = { $clusterSession.Process.AvailablePerfomance } },
                                                                                                                @{ Name = 'AvgBackCallTime';     Expression = { $clusterSession.Process.AvgBackCallTime } },
                                                                                                                @{ Name = 'AvgCallTime';         Expression = { $clusterSession.Process.AvgCallTime } },
                                                                                                                @{ Name = 'AvgDBCallTime';       Expression = { $clusterSession.Process.AvgDBCallTime } },
                                                                                                                @{ Name = 'AvgLockCallTime';     Expression = { $clusterSession.Process.AvgLockCallTime } },
                                                                                                                @{ Name = 'AvgServerCallTime';   Expression = { $clusterSession.Process.AvgServerCallTime } },
                                                                                                                @{ Name = 'AvgThreads';          Expression = { $clusterSession.Process.AvgThreads } },
                                                                                                                @{ Name = 'Capacity';            Expression = { $clusterSession.Process.Capacity } },
                                                                                                                @{ Name = 'Connections';         Expression = { $clusterSession.Process.Connections } },
                                                                                                                @{ Name = 'HostName';            Expression = { $clusterSession.Process.HostName } },
                                                                                                                @{ Name = 'IsEnable';            Expression = { $clusterSession.Process.IsEnable } },
                                                                                                                @{ Name = 'License';             Expression = { try { $clusterSession.Process.License.FullPresentation } catch { $null } } },
                                                                                                                @{ Name = 'MainPort';            Expression = { $clusterSession.Process.MainPort   } },
                                                                                                                @{ Name = 'MemoryExcessTime';    Expression = { $clusterSession.Process.MemoryExcessTime } },
                                                                                                                @{ Name = 'MemorySize';          Expression = { $clusterSession.Process.MemorySize } },
                                                                                                                @{ Name = 'PID';                 Expression = { $clusterSession.Process.PID } },
                                                                                                                @{ Name = 'Running';             Expression = { $clusterSession.Process.Running } },
                                                                                                                @{ Name = 'SelectionSize';       Expression = { $clusterSession.Process.SelectionSize } },
                                                                                                                @{ Name = 'StartedAt';           Expression = { $clusterSession.Process.Process.StartedAt } },
                                                                                                                @{ Name = 'Use';                 Expression = { $clusterSession.Process.Use } } } } },
                                                        @{ Name = 'Connection'; Expression = { if ( $ShowConnections -eq 'Everywhere') { 1 | Select-Object @{ Name = 'Application'; Expression = { $clusterSession.Application } },
                                                                                                                    @{ Name = 'blockedByLS'; Expression = { $clusterSession.blockedByLS } },
                                                                                                                    @{ Name = 'ConnectedAt'; Expression = { $clusterSession.ConnectedAt } },
                                                                                                                    @{ Name = 'ConnID';      Expression = { $clusterSession.ConnID } },
                                                                                                                    @{ Name = 'Host';        Expression = { $clusterSession.Host } },
                                                                                                                    @{ Name = 'InfoBase';    Expression = { @{ Descr = $clusterSession.InfoBase.Descr; Name = $clusterSession.InfoBase.Name } } },
                                                                                                                    @{ Name = 'SessionID';   Expression = { $clusterSession.SessionID } },
                                                                                                                    @{ Name = 'Process';     Expression = { if ( -not $NoWorkingProcesses ) { 1 | Select-Object   @{ Name = 'AvailablePerfomance'; Expression = { $clusterSession.Process.AvailablePerfomance } },
                                                                                                                                                                                @{ Name = 'AvgBackCallTime';     Expression = { $clusterSession.Process.AvgBackCallTime } },
                                                                                                                                                                                @{ Name = 'AvgCallTime';         Expression = { $clusterSession.Process.AvgCallTime } },
                                                                                                                                                                                @{ Name = 'AvgDBCallTime';       Expression = { $clusterSession.Process.AvgDBCallTime } },
                                                                                                                                                                                @{ Name = 'AvgLockCallTime';     Expression = { $clusterSession.Process.AvgLockCallTime } },
                                                                                                                                                                                @{ Name = 'AvgServerCallTime';   Expression = { $clusterSession.Process.AvgServerCallTime } },
                                                                                                                                                                                @{ Name = 'AvgThreads';          Expression = { $clusterSession.Process.AvgThreads } },
                                                                                                                                                                                @{ Name = 'Capacity';            Expression = { $clusterSession.Process.Capacity } },
                                                                                                                                                                                @{ Name = 'Connections';         Expression = { $clusterSession.Process.Connections } },
                                                                                                                                                                                @{ Name = 'HostName';            Expression = { $clusterSession.Process.HostName } },
                                                                                                                                                                                @{ Name = 'IsEnable';            Expression = { $clusterSession.Process.IsEnable } },
                                                                                                                                                                                @{ Name = 'License';             Expression = { try { $clusterSession.Process.License.FullPresentation } catch { $null } } },
                                                                                                                                                                                @{ Name = 'MainPort';            Expression = { $clusterSession.Process.MainPort   } },
                                                                                                                                                                                @{ Name = 'MemoryExcessTime';    Expression = { $clusterSession.Process.MemoryExcessTime } },
                                                                                                                                                                                @{ Name = 'MemorySize';          Expression = { $clusterSession.Process.MemorySize } },
                                                                                                                                                                                @{ Name = 'PID';                 Expression = { $clusterSession.Process.PID } },
                                                                                                                                                                                @{ Name = 'Running';             Expression = { $clusterSession.Process.Running } },
                                                                                                                                                                                @{ Name = 'SelectionSize';       Expression = { $clusterSession.Process.SelectionSize } },
                                                                                                                                                                                @{ Name = 'StartedAt';           Expression = { $clusterSession.Process.Process.StartedAt } },
                                                                                                                                                                                @{ Name = 'Use';                 Expression = { $clusterSession.Process.Use } } } } } } } }
                    }

                    Add-Member -InputObject $cls -Name Sessions -Value $objClSession -MemberType NoteProperty

                }

            }

            $obj.Clusters += $cls
            $result += $obj
                
        }

    }

    $result

    }

End {
    $connector = $null
    }

}



function Get-1CRegisteredApplicationClasses {
    
    param (
        
        # Имя компьютера
        [string]$ComputerName,

        # Номер версии платформы
        [ValidateScript({ $_ -match '\d\.\d\.\d+\.\d+' })]
        [string]$Version

    )
    
    $rv = @()

    $reg=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('ClassesRoot', $ComputerName)
            
    $regkey = $reg.OpenSubKey("CLSID")
    
    $subkeys = $regkey.GetSubKeyNames()

    foreach( $subkey in $subkeys )
    {
        if ( $subkey -notmatch "\{\w{8}\-(\w{4}\-){3}\w{12}\}" )
        {
            Continue
        }

        $curPath = "CLSID\\$subkey"

        $curSubKey = $reg.OpenSubKey("$curPath\\VersionIndependentProgID")

        if ( $curSubKey )
        {
            
            $defaultKeyValue = $curSubKey.GetValue("")

            if ( $defaultKeyValue -in @('V83.Application','V83C.Application') ) {
                
                $curLocalServer32Path = "$curPath\\LocalServer32"

                $curLocalServer32Key = $reg.OpenSubKey($curLocalServer32Path)

                $curImagePath = $curLocalServer32Key.getValue("")

                if ( Test-Path $curImagePath ) {
                    
                    $rv += 1 | Select-Object @{ Name = 'VersionIndependentProgID'; Expression = { $defaultKeyValue } },
                                             @{ Name = 'FilePath'; Expression = { $curImagePath } }, 
                                             @{ Name = 'DirectoryName'; Expression = { ([System.IO.FileInfo]$curImagePath).DirectoryName }}

                }

            } elseif ( $defaultKeyValue -in @("V83.COMConnector","V83.ServerAbout","V83.ServerAdminScope") ) {

                $curInprocServer32Path = "$curPath\\InprocServer32"

                $curInprocServer32Path = $reg.OpenSubKey($curInprocServer32Path)

                $curImagePath = $curInprocServer32Path.getValue("")

                if ( Test-Path $curImagePath ) {

                    $rv += 1 | Select-Object @{ Name = 'VersionIndependentProgID'; Expression = { $defaultKeyValue } },
                                             @{ Name = 'FilePath'; Expression = { $curImagePath } }, 
                                             @{ Name = 'DirectoryName'; Expression = { ([System.IO.FileInfo]$curImagePath).DirectoryName }}

                }

            }

        }
    
    }

    $rv

}


function Get-1CAppDirs {
    param (
        # Номер версии платформы
        [ValidateScript({ $_ -match '\d\.\d\.\d+\.\d+' })]
        [string]$Version
    )
    
    $possibleAppDirs = @("${env:ProgramFiles}\1cv8", "${env:ProgramFiles(x86)}\1cv8", "$ENV:USERPROFILE\AppData\Local\Programs\1cv8_x86", "$ENV:USERPROFILE\AppData\Local\Programs\1cv8")
    $possibleAppDirs = $possibleAppDirs | Where-Object { Test-Path $_ }

    if ( $Version ) {
        $pattern = $Version.Replace('.','\.')
    } else {
        $pattern = '\d\.\d\.\d+\.\d+'
    }

    $ENV:H1CAPPDIRS = @()
    ( Get-ChildItem -Directory -Path $possibleAppDirs | Where-Object { $_ -match $pattern } ).FullName | Select-Object -Unique | ForEach-Object { $ENV:H1CAPPDIRS += "$_\bin;" }
    ( Get-1CRegisteredApplicationClasses ).DirectoryName | Where-Object { $_ -match $pattern } | Select-Object -Unique | Where-Object { -not $ENV:H1CAPPDIRS.Contains($_) } | ForEach-Object { $ENV:H1CAPPDIRS += "$_;" }
    
    $ENV:H1CAPPDIRS

}


function Invoke-RAS {
    [CmdletBinding()]
    param (
        # Список параметров для ras.exe
        [Parameter(ValueFromRemainingArguments = $true)]
        [string]$ArgumentList = 'help'
    )

    if ( -not $ENV:H1CRASPATH -or ( $ENV:H1CRASVERSION -and $ENV:H1CRASPATH.Contains($ENV:H1CRASVERSION) )) {
        Set-RASversion -Verbose:$VerbosePreference | Out-Null
    }
    
    if ( $ENV:H1CRASPATH ) {
    
        Write-Verbose "Запуск '`"$ENV:H1CRASPATH`" $ArgumentList'"
        Start-Process -FilePath $ENV:H1CRASPATH -ArgumentList $ArgumentList -Wait -NoNewWindow -Verbose:$VerbosePreference
    
    } else {
        
        $errorMessage = 'Не удалось найти исполняемый файл клиента сервера удалённого администрирования ras.exe (заполните переменную окружения $ENV:H1CRASPATH)'
        
        if ( $ENV:H1CRACVERSION ) {

            $errorMessage += " для версии платформы $ENV:H1CRASVERSION"

        }
        
        Write-Error $errorMessage

    }

}


<#
.SYNOPSIS
    Создание нового экземпляра сервиса сервера администрирования 1С

.EXAMPLE
    $user = 'root'
    $pwd = Read-Host -AsSecureString
    New-RASservice 8.3.17.2171 $user $pwd
#>
function New-RASservice {
    
    [CmdletBinding()]
    param (

        # Версия платформы
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateScript({ $_ -match '\d\.\d\.\d+\.\d+' })]
        [string]$Version,

        [Parameter(Mandatory=$true, Position=1)]
        [string]$SrvUsrName,

        [Parameter(Mandatory=$true, Position=2)]
        [System.Security.SecureString]$SrvUsrPwd,

        # Адрес агента сервера
        [Parameter(Position=3)]
        [string]$AgentName = 'localhost',

        [ValidateScript({ $_ -match '\d{4,5}' })]
        [string]$CtrlPort = '1540',

        [ValidateScript({ $_ -match '\d{4,5}' })]
        [string]$ServicePort = '1545',

        # Путь к файлу ras.exe (только если требуется использовать определённый файл)
        [ValidateScript({ Test-Path $_  })]
        [string]$FilePath

    )
    
    if ( -not $FilePath ) {

        $possibleDirs = ( Get-1CAppDirs -Version:$Version ).split(';')

        $rasFileInfo = Get-ChildItem -File -Path $possibleDirs -Filter ras.exe | Select-Object -First 1

    } else {

        $rasFileInfo = Get-ChildItem -LiteralPath $FilePath

    }

    if ( $rasFileInfo ) {

        $ENV:H1CRASPATH = $rasFileInfo.FullName

        $serviceName = "1C:Enterprise $Version Remote Server"  

        $displayName = "Сервер администрирования 1С:Предприятие $Version"

        $binaryPathName = "`"${ENV:H1CRASPATH}`" cluster --service --port=${ServicePort} ${AgentName}:${CtrlPort}"

        $serviceWMIobject = Get-WmiObject -ClassName Win32_Service -Filter "Name='$serviceName'"
        
        if ( $serviceWMIobject ) {
            
            Write-Verbose "Остановка сервиса '$serviceName'"
            
            $serviceWMIobject.StopService() | Out-Null
            
            Write-Verbose "Удаление сервиса '$serviceName'"
            
            $serviceWMIobject.Delete() | Out-Null

        }

        $credential = New-Object System.Management.Automation.PSCredential ($SrvUsrName, $SrvUsrPwd)

        New-Service -Name $serviceName -DisplayName $displayName -Description $displayName -BinaryPathName $binaryPathName -StartupType Automatic -Credential $credential -Verbose:$VerbosePreference

        Start-Service -Name $serviceName -Verbose:$VerbosePreference

    } else {

        Write-Error "Не удалось найти ras.exe для версии платформы $Version"

    }

}


<#
.SYNOPSIS
    Установить используемую версию платформы для ras.exe на время сеанса
#>
function Set-RASversion {
    [CmdletBinding()]
    param (
        # Номер версии платформы
        [ValidateScript({ $_ -match '\d\.\d\.\d+\.\d+' })]
        [string]$Version
    )
    
    if ( $Version ) {

        $ENV:H1CRASVERSION = $Version
        Write-Verbose "Установлена версия платформы $ENV:H1CRASVERSION для использования ras.exe"

    } else {

        Invoke-RAC -DoNotInvokeEXE:$true -Verbose:$VerbosePreference 

        $rasPath = $ENV:H1CRACPATH.Replace('rac.exe', 'ras.exe')

        if ( Test-Path $rasPath ) {
            
            $ENV:H1CRASPATH = $rasPath
            
            [regex]$rxVersionNumber = '\d\.\d\.\d+\.\d+'

            $ENV:H1CRASVERSION = $rxVersionNumber.Matches( $rasPath ).Value | Sort-Object -Descending | Select-Object -First 1

        }

    }

    $ENV:H1CRASVERSION
    
}


<#
.SYNOPSIS
    Установить используемую версию платформы для rac.exe на время сеанса
#>
function Set-RACversion {
    [CmdletBinding()]
    param (
        # Номер версии платформы
        [ValidateScript({ $_ -match '\d\.\d\.\d+\.\d+' })]
        [string]$Version
    )
    
    if ( $Version ) {

        $ENV:H1CRACVERSION = $Version
        Write-Verbose "Установлена версия платформы $ENV:H1CRACVERSION для использования rac.exe"

    } else {

        Invoke-RAC -Mode 'help' -DoNotInvokeEXE:$true -Verbose:$VerbosePreference 

    }

    $ENV:H1CRACVERSION
    
}


<#
.SYNOPSIS
    Производит вызов rac.exe (предварительно производится поиск приложения)
    При обращении к серверу следует указывать имя сервера первым аргументом
#>
function Invoke-RAC {

    [CmdletBinding()]
    param (
        # Список параметров для rac.exe
        [Parameter(ValueFromRemainingArguments = $true)]
        [string]$ArgumentList = 'help',

        # Если не нужно вызывать rac.exe
        [switch]$DoNotInvokeEXE
    )

    if ( -not $ENV:H1CRACPATH -or ( $ENV:H1CRACVERSION -and -not $ENV:H1CRACPATH.Contains($Version) ) ) {

        # Поиск подходящей версии приложения

        $possibleAppDirs = if ( $ENV:H1CAPPDIRS ) { $ENV:H1CAPPDIRS.Split(';') } else { (Get-1CAppDirs).Split(';') }

        if ( $ENV:H1CRACVERSION ) {

            $possibleAppDirs = $possibleAppDirs | Where-Object { $_.Contains( $ENV:H1CRACVERSION ) }

        }

        $foundFiles = Get-ChildItem -File -Filter 'rac.exe' -Path $possibleAppDirs

        if ( $foundFiles ) {

            [regex]$rxVersionNumber = '\d\.\d\.\d+\.\d+'

            $maxVersion = $rxVersionNumber.Matches( $foundFiles.DirectoryName ).Value | Sort-Object -Descending | Select-Object -First 1

            if ( -not $ENV:H1CRACVERSION) {

                $ENV:H1CRACVERSION = $maxVersion
                Write-Verbose "Установлена версия платформы $ENV:H1CRACVERSION для использования rac.exe"

            }

            $ENV:H1CRACPATH = ( $foundFiles | Where-Object { $_.DirectoryName.Contains($maxVersion) } | Select-Object -First 1 ).FullName

        }

    }
    
    if ( -not $DoNotInvokeEXE -and $ENV:H1CRACPATH ) {
            
        Write-Verbose "Запуск '`"$ENV:H1CRACPATH`" $ArgumentList'"

        Start-Process -FilePath $ENV:H1CRACPATH -ArgumentList $ArgumentList -Wait -NoNewWindow -Verbose:$VerbosePreference 
    
    } elseif ( -not $DoNotInvokeEXE ) {
        
        $errorMessage = 'Не удалось найти исполняемый файл клиента сервера удалённого администрирования rac.exe (заполните переменную окружения $ENV:H1CRACPATH)'
        
        if ( $ENV:H1CRACVERSION ) {

            $errorMessage += " для версии платформы $ENV:H1CRACVERSION"

        }
        
        Write-Error $errorMessage

    }

}


function Get-1CHostData {

    param (

        [string]$ComputerName

    )
    
    @{
        'CIMAgentPs' = Get-CimInstance -ClassName Win32_Process -Filter "Name='ragent.exe'" -Property * -ComputerName:$ComputerName; 
        'CIMAgentSvc' = Get-CimInstance -ClassName Win32_Service -Filter "PathName like '%ragent.exe%'" -Property * -ComputerName:$ComputerName; 
        'CIMConfPs' = Get-CimInstance -ClassName Win32_Process -Filter "Name='crserver.exe'" -Property * -ComputerName:$ComputerName; 
        'CIMConfSvc' = Get-CimInstance -ClassName Win32_Service -Filter "PathName like '%crserver.exe%'" -Property * -ComputerName:$ComputerName; 
        'CIMMngrPs' = Get-CimInstance -ClassName Win32_Process -Filter "Name='rmngr.exe'" -Property * -ComputerName:$ComputerName; 
        'CIMMngrSvc' = Get-CimInstance -ClassName Win32_Service -Filter "PathName like '%rmngr.exe%'" -Property * -ComputerName:$ComputerName;  
        'CIMRasPs' = Get-CimInstance -ClassName Win32_Process -Filter "Name='ras.exe'" -Property * -ComputerName:$ComputerName; 
        'CIMRasSvc' = Get-CimInstance -ClassName Win32_Service -Filter "PathName like '%ras.exe%'" -Property * -ComputerName:$ComputerName;
    }

}


<#
.SYNOPSIS
    Удаляет сеанс с кластера 1с

.DESCRIPTION

.NOTES  
    Name: 1CHelper
    Author: yauhen.makei@gmail.com

.LINK  
    https://github.com/emakei/1CHelper.psm1

.EXAMPLE
    $data = Get-1CClusterData 1c-cluster.contoso.com -NoClusterAdmins -NoClusterManagers -NoWorkingServers -NoWorkingProcesses -NoClusterServices -ShowConnections None -ShowSessions Cluster -ShowLocks None -NoInfobases -NoAssignmentRules -User Example -Password Example
    Remove-1Csession -HostName $data.Clusters.HostName -MainPort $data.Clusters.MainPort -User Admin -Password Admin -SessionID 3076 -InfoBaseName TestDB -Verbose -NotCloseConnection

#>
function Remove-1CSession
{
[CmdletBinding()]
Param(
        # Адрес хоста для удаления сеанса
        [Parameter(Mandatory=$true)]
        [string]$HostName,
        # Порт хоста для удаления сеанса
        [Parameter(Mandatory=$true)]
        [int]$MainPort,
        # Имя админитратора кластера
        [string]$User="",
        # Пароль администратора кластера
        [Security.SecureString]$Password="",
        # Порт хоста для удаления сеанса
        [Parameter(Mandatory=$true)]
        [int]$SessionID,
        # Порт хоста для удаления сеанса
        [Parameter(Mandatory=$true)]
        [string]$InfoBaseName,
        # Принудительно закрыть соединение с информационной базой после удаления сеанса
        [switch]$CloseIbConnection=$false,
        # Имя админитратора информационной базы
        [string]$IbUser="",
        # Пароль администратора информационной базы
        [securestring]$IbPassword="",
        # Версия компоненты
        [ValidateSet(2, 3, 4)]
        [int]$Version=3
    )

Begin {
    $connector = New-Object -ComObject "v8$version.COMConnector"
    }

Process {

    try {
        Write-Verbose "Подключение к '$HostName'"
        $connection = $connector.ConnectAgent( $HostName )
        $abort = $false
    } catch {
        Write-Warning $_
        $abort = $true
    }
        
    if ( -not $abort ) {
            
        Write-Verbose "Подключен к `"$($connection.ConnectionString)`""

        $clusters = $connection.GetClusters()

        foreach ( $cluster in $clusters ) {
            
            if ( $cluster.HostName -ne $HostName -or $cluster.MainPort -ne $MainPort ) { continue }

            try {
                Write-Verbose "Аутентификация в кластере '$($cluster.HostName,':',$cluster.MainPort,' - ',$cluster.ClusterName)'"
                $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
                $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                $connection.Authenticate( $cluster, $User, $PlainPassword )
                $abort = $false
            } catch {
                Write-Warning $_
                continue
            }

            $sessions = $connection.GetSessions( $cluster )
                
            foreach ( $session in $sessions ) {
                    
                if ( $session.InfoBase.Name -ne $InfoBaseName -or $session.SessionID -ne $SessionID ) { continue }

                Write-Verbose "Удаление сеанса '$($session.SessionID,' - ''',$session.UserName,''' с компьютера : ',$session.Host)'"
                try {
                    $connection.TerminateSession( $cluster, $session )
                } catch {
                    Write-Warning $_
                    continue
                }
                
                if ( $CloseIbConnection -and $session.Connection ) {
                    try {
                        # подключаемся к рабочему процессу
                        Write-Verbose "Подключение к рабочему процессу '$($session.Process.HostName):$($session.Process.MainPort)'"
                        $server = $connector.ConnectWorkingProcess( "$($session.Process.HostName):$($session.Process.MainPort)" )
                        # проходим аутентификацию в информационной базе
                        Write-Verbose "Аутентификация пользователя инф. базы '$($IbUser)' в информационной базе '$($InfoBaseName)'"
                        $BSTRIbPassword = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($IbPassword)
                        $PlainIbPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTRIbPassword)
                        $server.AddAuthentication( $IbUser, $PlainIbPassword )
                        # отключаем соединение
                        $ibDesc = $server.CreateInfoBaseInfo()
                        $ibDesc.Name = $InfoBaseName
                        $ibConnections = $server.GetInfoBaseConnections( $ibDesc )
                        foreach ( $ibConnection in $ibConnections ) {
                            if ( $ibConnection.ConnID -ne $session.connection.ConnID ) { continue } 
                            # отключение соединения
                            Write-Verbose "Отключение соединения № '$($ibConnection.ConnID)' приложения '$($ibConnection.AppID)' c компьютера '$($ibConnection.HostName)'"
                            $server.Disconnect( $ibConnection )
                        }
                    } catch {
                        Write-Warning $_
                        continue
                    }

                }
                
            }

        }

    }
            
    }

End {
    $connector = $null
    }

}

<#
.SYNOPSIS
   Находит значения параметров в файле nethasp.ini

.DESCRIPTION

.NOTES  
    Name: 1CHelper
    Author: yauhen.makei@gmail.com

.LINK  
    https://github.com/emakei/1CHelper.psm1

.EXAMPLE
   Get-1CNetHaspIniStrings

.OUTPUTS
   Структура параметров
#>
function Get-1CNetHaspIniStrings
{
    
    $struct = @{}
    
    $pathToStarter = Find-1CEstart
    $pathToFile = $pathToStarter.Replace("common\1cestart.exe", "conf\nethasp.ini")

    if ( $pathToStarter ) {
        
        $content = Get-Content -Encoding UTF8 -LiteralPath $pathToFile
        $strings = $content | Where-Object { $_ -match "^\w" }
        $strings | ForEach-Object { $keyValue = $_.Split('='); $key = $keyValue[0].Replace(" ",""); $value = $keyValue[1].Replace(" ",""); $value = $value.Replace(';',''); $struct[$key] = $value.Split(',') }

    }

    $struct

}

<#
.SYNOPSIS
   Поиск максимальной версии приложения

.DESCRIPTION
   Поиск максимальной версии приложения (не ниже 8.3)

.NOTES  
    Name: 1CHelper
    Author: yauhen.makei@gmail.com

.LINK  
    https://github.com/emakei/1CHelper.psm1

.EXAMPLE
   Find-1CApplicationForExportImport

.OUTPUTS
   NULL или строку с путем установки приложения
#>
function Find-1CApplicationForExportImport
{
    [CmdletBinding()]
    Param(
        
        # Имя компьютера для поиска версии
        [string]$ComputerName,

        # Вывести все результаты
        [switch]$AllInstances
    )

    $installationPath = if( $AllInstances ) { @() } else { $null }

    $pvs = 0

    $UninstallPaths = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")
   
    ForEach($UninstallKey in $UninstallPaths) {
        
         Try {
             $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
         } Catch {
             Write-Error $_
             Continue
         }
 
         $regkey = $reg.OpenSubKey($UninstallKey)

         If(-not $regkey) {
             Write-Warning "Не найдены ключи в: HKLM:\\$UninstallKey"
         }

         $subkeys = $regkey.GetSubKeyNames()
        
         foreach($key in $subkeys){
 
             $thisKey = $UninstallKey + "\\" + $key
 
             $thisSubKey = $reg.OpenSubKey($thisKey)

             Try {
                 $displayVersion = $thisSubKey.getValue("DisplayVersion").Split('.')
             } Catch {
                 Continue
             }
            
             if ( $displayVersion.Count -ne 4) { continue }
             if ( -not ($thisSubKey.getValue("Publisher") -in @("1C","1С") `
                     -and $displayVersion[0] -eq 8 `
                     -and $displayVersion[1] -gt 2 ) ) { continue }
             $tmpPath = $thisSubkey.getValue("InstallLocation")
             if (-not $tmpPath.EndsWith('\')) {
                 $tmpPath += '\' + 'bin\1cv8.exe'
             } else {
                 $tmpPath += 'bin\1cv8.exe'
             }
             Try {
                 $tmpPVS = [double]$displayVersion[1] * [Math]::Pow(10, 6) + [double]$displayVersion[2] * [Math]::Pow(10, 5) + [double]$displayVersion[3]
             } Catch {
                 Continue
             }
             if ( $tmpPVS -gt $pvs -and ( Test-Path -LiteralPath $tmpPath) ) {
                 $pvs = $tmpPVS
                 if ( $AllInstances ) {
                    $installationPath += Get-Item $tmpPath
                 }
                 else {
                    $installationPath = $tmpPath
                 }
             }
 
         }

         $reg.Close() 

     }

     # https://docs.microsoft.com/en-us/windows/win32/shell/app-registration
     # В 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths' нету 1С
     # как и в 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths'

     # Поиск в доступных COM-классах

     if ( -not $installationPath -or $AllInstances )
     {
        Try {
            
            $reg=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('ClassesRoot', $ComputerName)
            
            $regkey = $reg.OpenSubKey("CLSID")
            
            $subkeys = $regkey.GetSubKeyNames()

            foreach( $subkey in $subkeys )
            {
                if ( $subkey -notmatch "\{\w{8}\-(\w{4}\-){3}\w{12}\}" )
                {
                    Continue
                }

                $curPath = "CLSID\\$subkey"

                $curSubKey = $reg.OpenSubKey("$curPath\\VersionIndependentProgID")

                if ( $curSubKey )
                {
                    
                    $defaultKeyValue = $curSubKey.GetValue("")

                    if ( $defaultKeyValue -eq 'V83.Application') {
                        
                        $curLocalServer32Path = "$curPath\\LocalServer32"

                        $curLocalServer32Key = $reg.OpenSubKey($curLocalServer32Path)

                        $curImagePath = $curLocalServer32Key.getValue("")

                        if ( Test-Path $curImagePath ) {
                            if ( $AllInstances ) {
                                $installationPath += Get-Item $curImagePath
                            } else {
                                $installationPath = $curImagePath
                            }
                        }

                    } elseif ( $defaultKeyValue -eq "V83.COMConnector" ) {
  
                        $curInprocServer32Path = "$curPath\\InprocServer32"

                        $curInprocServer32Path = $reg.OpenSubKey($curInprocServer32Path)

                        $curImagePath = $curInprocServer32Path.getValue("").Replace('comcntr.dll','1cv8.exe')

                        if ( Test-Path $curImagePath ) {

                            if ( $AllInstances ) {
                                $installationPath += Get-Item $curImagePath
                            } else {
                                $installationPath = $curImagePath
                            }

                        }

                    }

                }

            }

        } Catch {
            
            Write-Error $_
            Continue

        }
    }

    $installationPath
}

<#
.SYNOPSIS
   Возвращает путь к каталогу с библиотекой hsmon.dll

.DESCRIPTION

.NOTES  
    Name: 1CHelper
    Author: yauhen.makei@gmail.com

.LINK  
    https://github.com/emakei/1CHelper.psm1

.EXAMPLE
   Get-NetHaspDirectoryPath

.OUTPUTS
   Путь к каталогу с библиотекой hsmon.dll
#>
function Get-NetHaspDirectoryPath
{  
    (Get-Module 1CHelper).Path.TrimEnd('1CHelper.psm1') + "hasp"
}

<#
.SYNOPSIS
   Возвращает путь к файлу nethasp.ini

.DESCRIPTION

.NOTES  
    Name: 1CHelper
    Author: yauhen.makei@gmail.com

.LINK  
    https://github.com/emakei/1CHelper.psm1

.EXAMPLE
   Get-NetHaspIniFilePath

.OUTPUTS
   Путь к к файлу nethasp.ini
#>
function Get-NetHaspIniFilePath
{  
    $pathToStarter = Find-1CEstart
    $pathToStarter.Replace("common\1cestart.exe", "conf\nethasp.ini")
}

<#
.SYNOPSIS
   Возвращает результат выполнения запроса к серверу SQL

.DESCRIPTION

.NOTES  
    Name: 1CHelper
    Author: yauhen.makei@gmail.com

.LINK  
    https://github.com/emakei/1CHelper.psm1


.EXAMPLE
    Invoke-SqlQuery -Server test.contoso.com -Database test -user admin -password admin -Text 'select @@version'

.EXAMPLE
    Invoke-SqlQuery -Server test.contoso.com -Database test -user admin -password admin -Data 'Свободно в tempdb' -Verbose

.EXAMPLE
    Invoke-SqlQuery -Server test.contoso.com -user admin -password admin -Data 'Наибольшая нагрузка на CPU' -Verbose

#>
function Invoke-SqlQuery
{
    [CmdletBinding()]
Param (
    [string]$Server='local',
    
    [string]$Database,
    
    [Parameter(Mandatory=$true)]
    [string]$user,
    
    [Parameter(Mandatory=$true)]
    [Security.SecureString]$password,

    [Parameter(Mandatory=$false)]
    [ValidateScript({ -not $Text })]
    [ValidateSet(
        ,'Базы создающие нагрузку на диск'
        ,'Длительные транзакции'
        ,'Запросы с высокими издержками на ввод/вывод'
        ,'Использование кешей по базам данных сервера СУБД'
        ,'Использование кешей сервера СУБД'
        ,'Используемые протоколы подключения'
        ,'Количество ожидающих исполнителей, которые ждут исполнения задания'
        ,'Нагрузка на CPU по базам'
        ,'Наиболее часто выполняемые запросы'
        ,'Наибольшая нагрузка на CPU'
        ,'Определение объема пространства, используемого внутренними объектами'
        ,'Определение объема пространства, используемого пользовательскими объектами'
        ,'Определение объема пространства, используемого хранилищем версий'
        ,'Определение объема свободного пространства в tempdb'
        ,'Оценить наличие и величину ожидания ввода-вывода'
        ,'Оценить наличие и величину ожидания при синхронизации потоков выполнения'
        ,'Получение пространства, занимаемого внутренними объектами в текущем сеансе как для выполнения, так и для выполнения задач'
        ,'Получение пространства, занимаемого внутренними объектами во всех текущих выполняемых задачах в каждом сеансе'
        ,'Проверка включена ли опция оптимизации tempdb по памяти'
        ,'Проверка текущего размера и параметра авторасширения tempdb'
        ,'Проверка фрагментации индекса хранилища столбцов'
        ,'Свободно в tempdb'
        ,'Список длительных транзакций'
        ,'Текущая статистика по задержкам'
        ,'Топ запросов, создающих нагрузку на CPU на сервере СУБД за последний час'
        ,'Фрагментация и плотность страниц индекса хранилища строк'
                )]
    [string]$Data,
    
    [ValidateScript({ -not $Data })]
    [string]$Text
    )

    switch ( $Data ) {
        'Количество ожидающих исполнителей, которые ждут исполнения задания'
        {
            $sql = @'
            select max([runnable_tasks_count]) as [runnable_tasks_count]
            from sys.dm_os_schedulers
            where scheduler_id<255;
'@
        }
        'Оценить наличие и величину ожидания ввода-вывода'
        {
            $sql = @'
            with waits
            as
            (
            select
                   wait_type,
                   wait_time_ms,
                   waiting_tasks_count
            from sys.dm_os_wait_stats
            )
            select
                   waits.wait_type Wait_type,
                   waits.waiting_tasks_count Waiting_tasks,
                   waits.wait_time_ms Wait_time,
                   100 * waits.wait_time_ms / Totals.Total Percentage
            from waits
            inner join
                         (
                         select
                                sum (waits.wait_time_ms) Total
                         from waits
                         ) Totals
            on 1=1
            where waits.wait_type = N'IO'
'@
        }
        'Оценить наличие и величину ожидания при синхронизации потоков выполнения'
        {
            $sql = @'
            with waits
            as
            (
            select
                   wait_type,
                   wait_time_ms,
                   waiting_tasks_count
            from sys.dm_os_wait_stats
            )
            select
                   waits.wait_type Wait_type,
                   waits.waiting_tasks_count Waiting_tasks,
                   waits.wait_time_ms Wait_time,
                   100 * waits.wait_time_ms / Totals.Total Percentage
            from waits
            inner join
                         (
                         select
                                sum (waits.wait_time_ms) Total
                         from waits
                         ) Totals
            on 1=1
            where waits.wait_type = N'CXPACKET'
'@
        }
        'Получение пространства, занимаемого внутренними объектами во всех текущих выполняемых задачах в каждом сеансе'
        {
            $sql = @'
SELECT session_id,
  SUM(internal_objects_alloc_page_count) AS task_internal_objects_alloc_page_count,
  SUM(internal_objects_dealloc_page_count) AS task_internal_objects_dealloc_page_count
FROM sys.dm_db_task_space_usage
GROUP BY session_id;
'@
        }
        'Получение пространства, занимаемого внутренними объектами в текущем сеансе как для выполнения, так и для выполнения задач'
        {
            $sql = @'
SELECT R2.session_id,
  R1.internal_objects_alloc_page_count
  + SUM(R2.internal_objects_alloc_page_count) AS session_internal_objects_alloc_page_count,
  R1.internal_objects_dealloc_page_count
  + SUM(R2.internal_objects_dealloc_page_count) AS session_internal_objects_dealloc_page_count
FROM sys.dm_db_session_space_usage AS R1
INNER JOIN sys.dm_db_task_space_usage AS R2 ON R1.session_id = R2.session_id
GROUP BY R2.session_id, R1.internal_objects_alloc_page_count,
  R1.internal_objects_dealloc_page_count;
'@
        }
        'Определение объема свободного пространства в tempdb'
        {
            $sql = @'
SELECT SUM(unallocated_extent_page_count) AS [free pages],
  (SUM(unallocated_extent_page_count)*1.0/128) AS [free space in MB]
FROM tempdb.sys.dm_db_file_space_usage;
'@
        }
        'Определение объема пространства, используемого хранилищем версий'
        {
            $sql = @'
SELECT SUM(version_store_reserved_page_count) AS [version store pages used],
  (SUM(version_store_reserved_page_count)*1.0/128) AS [version store space in MB]
FROM tempdb.sys.dm_db_file_space_usage;
'@
        }
        'Определение объема пространства, используемого внутренними объектами'
        {
            $sql = @'
SELECT SUM(internal_object_reserved_page_count) AS [internal object pages used],
  (SUM(internal_object_reserved_page_count)*1.0/128) AS [internal object space in MB]
FROM tempdb.sys.dm_db_file_space_usage;
'@
        }
        'Определение объема пространства, используемого пользовательскими объектами'
        {
            $sql = @'
SELECT SUM(user_object_reserved_page_count) AS [user object pages used],
  (SUM(user_object_reserved_page_count)*1.0/128) AS [user object space in MB]
FROM tempdb.sys.dm_db_file_space_usage;
'@
        }
        'Проверка включена ли опция оптимизации tempdb по памяти'
        {
            $sql = @'
SELECT SERVERPROPERTY('IsTempdbMetadataMemoryOptimized')
'@
        }
        'Проверка текущего размера и параметра авторасширения tempdb'
        {
            $sql = @'
SELECT name AS FileName,
    size*1.0/128 AS FileSizeInMB,
    CASE max_size
        WHEN 0 THEN 'Autogrowth is off.'
        WHEN -1 THEN 'Autogrowth is on.'
        ELSE 'Log file grows to a maximum size of 2 TB.'
    END,
    growth AS 'GrowthValue',
    'GrowthIncrement' =
        CASE
            WHEN growth = 0 THEN 'Size is fixed.'
            WHEN growth > 0 AND is_percent_growth = 0
                THEN 'Growth value is in 8-KB pages.'
            ELSE 'Growth value is a percentage.'
        END
FROM tempdb.sys.database_files;
'@
        }
        'Проверка фрагментации индекса хранилища столбцов '
        {
            $sql = @'
SELECT OBJECT_SCHEMA_NAME(i.object_id) AS schema_name,
       OBJECT_NAME(i.object_id) AS object_name,
       i.name AS index_name,
       i.type_desc AS index_type,
       100.0 * (ISNULL(SUM(rgs.deleted_rows), 0)) / NULLIF(SUM(rgs.total_rows), 0) AS avg_fragmentation_in_percent
FROM sys.indexes AS i
INNER JOIN sys.dm_db_column_store_row_group_physical_stats AS rgs
ON i.object_id = rgs.object_id
   AND
   i.index_id = rgs.index_id
WHERE rgs.state_desc = 'COMPRESSED'
GROUP BY i.object_id, i.index_id, i.name, i.type_desc
ORDER BY schema_name, object_name, index_name, index_type;
'@
        }
        'Фрагментация и плотность страниц индекса хранилища строк'
        {
            $sql = @'
SELECT OBJECT_SCHEMA_NAME(ips.object_id) AS schema_name,
       OBJECT_NAME(ips.object_id) AS object_name,
       i.name AS index_name,
       i.type_desc AS index_type,
       ips.avg_fragmentation_in_percent,
       ips.avg_page_space_used_in_percent,
       ips.page_count,
       ips.alloc_unit_type_desc
FROM sys.dm_db_index_physical_stats(DB_ID(), default, default, default, 'SAMPLED') AS ips
INNER JOIN sys.indexes AS i 
ON ips.object_id = i.object_id
   AND
   ips.index_id = i.index_id
ORDER BY page_count DESC
'@
        }
        'Нагрузка на CPU по базам'
        {
            $sql = @'
WITH DB_CPU_Stats
AS
(SELECT DatabaseID, DB_Name(DatabaseID) AS [DatabaseName], SUM(total_worker_time) AS [CPU_Time_Ms]
FROM sys.dm_exec_query_stats AS qs
CROSS APPLY (SELECT CONVERT(int, value) AS [DatabaseID]
                FROM sys.dm_exec_plan_attributes(qs.plan_handle)
                WHERE attribute = N'dbid') AS F_DB
GROUP BY DatabaseID)
SELECT ROW_NUMBER() OVER(ORDER BY [CPU_Time_Ms] DESC) AS [row_num],
    DatabaseName, [CPU_Time_Ms],
    CAST([CPU_Time_Ms] * 1.0 / SUM([CPU_Time_Ms]) OVER() * 100.0 AS DECIMAL(5,2)) AS [CPUPercent]
FROM DB_CPU_Stats
WHERE DatabaseID > 4 -- system databases
AND DatabaseID <> 32767 -- ResourceDB
ORDER BY row_num OPTION (RECOMPILE)
'@
        }
        'Наибольшая нагрузка на CPU'
        {
            $sql = @'
SELECT TOP 10
    [Average CPU used] = total_worker_time / qs.execution_count
    , [Total CPU used] = total_worker_time
    , [Execution count] = qs.execution_count
    , [Individual Query] = SUBSTRING(qt.text, qs.statement_start_offset / 2 + 1, 
                                        (CASE 
                                            WHEN qs.statement_end_offset = -1 
                                                THEN LEN(CONVERT(NVARCHAR(MAX), qt.text)) * 2 
                                            ELSE qs.statement_end_offset 
                                        END - qs.statement_start_offset) / 2 + 1)
    , [Parent Query] = qt.text
    , DatabaseName = DB_NAME(qt.dbid)
FROM sys.dm_exec_query_stats qs
CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) as qt
ORDER BY [Average CPU used] DESC
'@
        }
        'Топ запросов, создающих нагрузку на CPU на сервере СУБД за последний час'
        {
            $sql = @'
SELECT
SUM(qs.max_elapsed_time) as elapsed_time,
SUM(qs.total_worker_time) as worker_time
INTO T1 FROM (
    SELECT TOP 100000
    *
    FROM sys.dm_exec_query_stats qs
    WHERE qs.last_execution_time > (CURRENT_TIMESTAMP - '01:00:00.000')
    ORDER BY qs.last_execution_time DESC
) as qs
;
SELECT TOP 10000
(qs.max_elapsed_time) as elapsed_time,
(qs.total_worker_time) as worker_time,
qp.query_plan,
st.text,
dtb.name,
qs.*,
st.dbid
INTO T2
FROM sys.dm_exec_query_stats qs
CROSS APPLY sys.dm_exec_query_plan(qs.plan_handle) qp
CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) st
LEFT OUTER JOIN sys.databases as dtb on st.dbid = dtb.database_id
WHERE qs.last_execution_time > (CURRENT_TIMESTAMP - '01:00:00.000')
ORDER BY qs.last_execution_time DESC
;
SELECT TOP 100
(T2.elapsed_time*100/T1.elapsed_time) as percent_elapsed_time,
(T2.worker_time*100/T1.worker_time) as percent_worker_time,
T2.*
FROM
T2 as T2
INNER JOIN T1 as T1
ON 1=1
ORDER BY T2.worker_time DESC
;
DROP TABLE T2
;
DROP TABLE T1
;
'@
        }
        'Наибольшая нагрузка на CPU'
        {
            $sql = @'
SELECT TOP 10
    [Average CPU used] = total_worker_time / qs.execution_count
    , [Total CPU used] = total_worker_time
    , [Execution count] = qs.execution_count
    , [Individual Query] = SUBSTRING(qt.text, qs.statement_start_offset / 2 + 1, 
                                        (CASE 
                                        WHEN qs.statement_end_offset = -1 
                                            THEN LEN(CONVERT(NVARCHAR(MAX), qt.text)) * 2 
                                        ELSE qs.statement_end_offset 
                                        END - qs.statement_start_offset) / 2 + 1)
    , [Parent Query] = qt.text
    , DatabaseName = DB_NAME(qt.dbid)
FROM sys.dm_exec_query_stats qs
CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) as qt
ORDER BY [Average CPU used] DESC
'@
        }
        'Список длительных транзакций'
        {
            $sql = @'
select
  transaction_id, *
from sys.dm_tran_active_snapshot_database_transactions
order by elapsed_time_seconds desc
'@
        }
        'Свободно в tempdb'
        {
            $sql = @'
select 
  sum(unallocated_extent_page_count) as [free pages]
  , (sum(unallocated_extent_page_count)*1.0/128) as [free space in MB]
from sys.dm_db_file_space_usage
'@
        }
        'Использование кешей по базам данных сервера СУБД'
        {
            $sql = @'
select db_name(database_id) as [Database name], count(row_count)*8.00/1024.00 as MB, count(row_count)*8.00/1024.00/1024.00 as GB
from sys.dm_os_buffer_descriptors
group by database_id
order by MB desc
'@
        }
        'Использование кешей сервера СУБД'
        {
            $sql = @'
select top(100)
[type]
, sum(pages_kb) as [SPA Mem, Kb]
from sys.dm_os_memory_clerks t
group by [type]
order by sum(pages_kb) desc
'@
        }
        'Запросы с высокими издержками на ввод/вывод'
        {
            $sql = @'
select top 100
  [Average IO] = (total_logical_reads + total_logical_writes) / qs.execution_count
, [Total IO] = (total_logical_reads + total_logical_writes)
, [Execution count] = qs.execution_count
, [Individual Query] = SUBSTRING(qt.text, qs.statement_start_offset/2 + 1, (case when qs.statement_end_offset = -1 then len(convert(nvarchar(max), qt.text)) * 2 else qs.statement_end_offset end - qs.statement_start_offset)/2)
, [Parent Query] = qt.text
, [Database name] = db_name(qt.dbid)
from sys.dm_exec_query_stats qs
cross apply sys.dm_exec_sql_text(qs.sql_handle) as qt
order by [Average IO] desc
'@
        }
        'Длительные транзакции'
        {
            $sql = @'
DECLARE @curr_date as DATETIME
SET @curr_date = GETDATE()
SELECT
    -- SESSION_TRAN.*,
    SESSION_TRAN.session_id AS ConnectionID, -- "Соединение с СУБД" в консоли кластера 1С
    -- TRAN_INFO.*,
    TRAN_INFO.transaction_begin_time,
    DateDiff(MINUTE, TRAN_INFO.transaction_begin_time, @curr_date) AS Duration, -- Длительность в минутах
    TRAN_INFO.transaction_type,
    -- 1 = транзакция чтения-записи;
    -- 2 = транзакция только для чтения;
    -- 3 = системная транзакция;
    -- 4 = распределенная транзакция.
    TRAN_INFO.transaction_state,
    -- 0 = транзакция ещё не была полностью инициализирована;
    -- 1 = транзакция была инициализирована, но ещё не началась;
    -- 2 = транзакция активна;
    -- 3 = транзакция закончилась;
    -- 4 = фиксирующий процесс был инициализирован на распределенной транзакции. Предназначено только для распределенных транзакций. Распределенная транзакция все еще активна, на дальнейшая обработка не может иметь место;
    -- 5 = транзакция находится в готовом состоянии и ожидает разрешения;
    -- 6 = транзакция зафиксирована;
    -- 7 = проводится откат транзакции;
    -- 8 = откат транзакции был выполнен.
    -- CONN_INFO.*,
    CONN_INFO.connect_time,
    CONN_INFO.num_reads,
    CONN_INFO.num_writes,
    CONN_INFO.last_read,
    CONN_INFO.last_write,
    CONN_INFO.client_net_address,
    CONN_INFO.most_recent_sql_handle,
    -- SQL_TEXT.*,
    SQL_TEXT.dbid,
    db_name(SQL_TEXT.dbid) as IB_NAME,
    SQL_TEXT.text,
    -- QUERIES_INFO.*,
    QUERIES_INFO.start_time,
    QUERIES_INFO.status,
    QUERIES_INFO.command,
    QUERIES_INFO.wait_type,
    QUERIES_INFO.wait_time,
    -- PLAN_INFO.*,
    PLAN_INFO.query_plan
FROM sys.dm_tran_session_transactions AS SESSION_TRAN
    JOIN sys.dm_tran_active_transactions as TRAN_INFO
    ON SESSION_TRAN.transaction_id = TRAN_INFO.transaction_id
    LEFT JOIN sys.dm_exec_connections AS CONN_INFO
    ON SESSION_TRAN.session_id = CONN_INFO.session_id
CROSS APPLY sys.dm_exec_sql_text(CONN_INFO.most_recent_sql_handle) AS SQL_TEXT
    LEFT JOIN sys.dm_exec_requests AS QUERIES_INFO
    ON SESSION_TRAN.session_id = QUERIES_INFO.session_id
    LEFT JOIN (
    SELECT
    VL_SESSION_TRAN.session_id AS session_id,
    VL_PLAN_INFO.query_plan AS query_plan
    FROM sys.dm_tran_session_transactions AS VL_SESSION_TRAN
    INNER JOIN sys.dm_exec_requests AS VL_QUERIES_INFO
    ON VL_SESSION_TRAN.session_id = VL_QUERIES_INFO.session_id
    CROSS APPLY sys.dm_exec_text_query_plan(VL_QUERIES_INFO.plan_handle, VL_QUERIES_INFO.statement_start_offset, VL_QUERIES_INFO.statement_end_offset) AS VL_PLAN_INFO) AS PLAN_INFO
    ON SESSION_TRAN.session_id = PLAN_INFO.session_id
ORDER BY transaction_begin_time ASC
'@
        }
        'Базы создающие нагрузку на диск'
        {
            $sql = @'
with
DB_Disk_Reads_Stats
as
(
    select DatabaseID, db_name(DatabaseID) as DatabaseName, sum(total_physical_reads) as physical_reads
    from sys.dm_exec_query_stats qs
cross apply (select convert(int, value) as DatabaseID
    from sys.dm_exec_plan_attributes(qs.plan_handle)
    where attribute = N'dbid') as F_DB
    group by DatabaseID
)
select ROW_NUMBER() OVER(ORDER BY [physical_reads] desc) as [row_num],
  DatabaseName, [physical_reads],
  CAST([physical_reads]*1.0/sum([physical_reads]) OVER() * 100.0 AS decimal(5, 2)) as [Physical_Reads_Percent]
from DB_Disk_Reads_Stats
where DatabaseID > 4 -- system databases
  and DatabaseID <> 32767
-- ResourceDB
order by row_num
OPTION
  (RECOMPILE)
'@
        }
        'Текущая статистика по задержкам'
        {
            $sql = @'
select top 100
  [Wait type] = wait_type,
  [Wait time (s)] = wait_time_ms / 1000,
  [% waiting] = convert(decimal(12,2), wait_time_ms * 100.0 / sum(wait_time_ms) OVER())
from sys.dm_os_wait_stats
where wait_type not like '%SLEEP%'
order by wait_time_ms desc
'@
        }
        'Наиболее часто выполняемые запросы'
        {
            $sql = @'
SELECT TOP 10
    [Execution count] = execution_count
    , [Individual Query] = SUBSTRING(qt.text, qs.statement_start_offset / 2 + 1,
                                    (CASE 
                                        WHEN qs.statement_end_offset = -1
                                            THEN LEN(CONVERT(NVARCHAR(MAX), qt.text))*2
                                        ELSE qs.statement_end_offset
                                    END - qs.statement_start_offset) / 2 + 1)
    , [Parent Query] = qt.text
    , [Database Name] = db_name(qt.dbid)
FROM sys.dm_exec_query_stats qs
CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) as qt
ORDER BY [Execution count] DESC
'@
        }
        'Используемые протоколы подключения'
        {
            $sql = @'
select program_name,net_transport
from sys.dm_exec_sessions as t1
left join sys.dm_exec_connections AS t2 ON t1.session_id=t2.session_id
where not t1.program_name is null
'@
        }
        default # не "Custom", т.к. проверяется параметр "Data"
        {
            $sql = $Text
        }
    }

    Write-Verbose "Текст запроса`n`n$sql`n`n"

    $connectionString = "Server=$Server"
    if ($Database)
    {
        $connectionString += ";Database=$Database"
    }

    Write-Verbose "Подключение к '$connectionString'"

    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList "$connectionString;Uid=$user;Pwd=$PlainPassword"
    $connection.Open()
    
    if ($connection.State -eq 'Open')
    {
        $command = New-Object -TypeName System.Data.SqlClient.SqlCommand $sql, $connection -ErrorAction Stop

        $adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
        $table = New-Object -TypeName System.Data.DataTable

        $rows = $adapter.Fill($table)
        
        Write-Verbose "Получено $rows строк(-а)"

        $connection.Close()
        $connection.Dispose()

        $table
    }
}

#region Zabbix
# https://github.com/zbx-sadman

#
#  Select object with Property that equal Value if its given or with Any Property in another case
#
Function PropertyEqualOrAny {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject,
      [PSObject]$Property,
      [PSObject]$Value
   );
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
         # IsNullorEmpty used because !$Value give a erong result with $Value = 0 (True).
         # But 0 may be right ID  
         If (($Object.$Property -Eq $Value) -Or ([string]::IsNullorEmpty($Value))) { $Object }
      }
   } 
}

#
#  Prepare string to using with Zabbix 
#
#Function PrepareTo-Zabbix {
Function Format-ToZabbix {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject,
      [String]$ErrorCode,
      [Switch]$NoEscape,
      [Switch]$JSONCompatible
   );
   Begin {
      # Add here more symbols to escaping if you need
      $EscapedSymbols = @('\', '"');
      $UnixEpoch = Get-Date -Date "01/01/1970";
   }
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
         If ($Null -Eq $Object) {
           # Put empty string or $ErrorCode to output  
           If ($ErrorCode) { $ErrorCode } Else { "" }
           Continue;
         }
         # Need add doublequote around string for other objects when JSON compatible output requested?
         $DoQuote = $False;
         Switch (($Object.GetType()).FullName) {
            'System.Boolean'  { $Object = [int]$Object; }
            'System.DateTime' { $Object = (New-TimeSpan -Start $UnixEpoch -End $Object).TotalSeconds; }
            Default           { $DoQuote = $True; }
         }
         # Normalize String object
         $Object = $( If ($JSONCompatible) { $Object.ToString() } else { $Object | Out-String }).Trim();
         
         If (!$NoEscape) { 
            ForEach ($Symbol in $EscapedSymbols) { 
               $Object = $Object.Replace($Symbol, "\$Symbol");
            }
         }

         # Doublequote object if adherence to JSON standart requested
         If ($JSONCompatible -And $DoQuote) { 
            "`"$Object`"";
         } else {
            $Object;
         }
      }
   }
}

#
#  Make & return JSON, due PoSh 2.0 haven't Covert-ToJSON
#
#Function Make-JSON {
Function Get-NetHaspJSON {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject, 
      [array]$ObjectProperties, 
      [Switch]$Pretty
   ); 
   Begin   {
      [String]$Result = "";
      # Pretty json contain spaces, tabs and new-lines
      If ($Pretty) { $CRLF = "`n"; $Tab = "    "; $Space = " "; } Else { $CRLF = $Tab = $Space = ""; }
      # Init JSON-string $InObject
      $Result += "{$CRLF$Space`"data`":[$CRLF";
      # Take each Item from $InObject, get Properties that equal $ObjectProperties items and make JSON from its
      $itFirstObject = $True;
   } 
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) {
         # Skip object when its $Null
         If ($Null -Eq $Object) { Continue; }

         If (-Not $itFirstObject) { $Result += ",$CRLF"; }
         $itFirstObject=$False;
         $Result += "$Tab$Tab{$Space"; 
         $itFirstProperty = $True;
         # Process properties. No comma printed after last item
         ForEach ($Property in $ObjectProperties) {
            If (-Not $itFirstProperty) { $Result += ",$Space" }
            $itFirstProperty = $False;
            $Result += "`"{#$Property}`":$Space$(Format-ToZabbix <#PrepareTo-Zabbix#> -InputObject $Object.$Property -JSONCompatible)";
         }
         # No comma printed after last string
         $Result += "$Space}";
      }
   }
   End {
      # Finalize and return JSON
      "$Result$CRLF$Tab]$CRLF}";
   }
}

#
#  Return value of object's metric defined by key-chain from $Keys Array
#
Function Get-Metric { 
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject, 
      [Array]$Keys
   ); 
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
        If ($Null -Eq $Object) { Continue; }
        # Expand all metrics related to keys contained in array step by step
        ForEach ($Key in $Keys) {              
           If ($Key) {
              $Object = Select-Object -InputObject $Object -ExpandProperty $Key -ErrorAction SilentlyContinue;
              If ($Error) { Break; }
           }
        }
        $Object;
      }
   }
}

#Function Compile-WrapperDLL() {
Function New-WrapperDLL() {
   $WrapperSourceCode = 
@"
   using System;
   using System.Runtime.InteropServices;
   using System.Text;
   
   namespace HASP { 
      public class Monitor { 
         [DllImport(`"$($HSMON_LIB_FILE)`", CharSet = CharSet.Ansi,EntryPoint=`"mightyfunc`", CallingConvention=CallingConvention.Cdecl)]
         // String type used for request due .NET do auto conversion to Ansi char* with marshaliing procedure;
         // Byte[] type used for response due .NET char* is 2-byte, but mightyfunc() need to 1-byte Ansi char;
         // Int type used for responseBufferSize due .NET GetString() operate with [int] params. So, response lenght must be Int32 sized
         extern static unsafe void mightyfunc(string request, byte[] response, int *responseBufferSize);
     
         public Monitor() {}
      
         public static unsafe string doCmd(string request) {
            int responseBufferSize = 10240, responseLenght = 0;
            byte[] response = new byte[responseBufferSize];
            string returnValue = `"`";
            mightyfunc(request, response, &responseBufferSize);
            while (response[responseLenght++] != '\0') 
            returnValue = System.Text.Encoding.UTF8.GetString(response, 0, responseLenght);
            return returnValue;
         }
      } 
   }
"@

   $CompilerParameters = New-Object -TypeName System.CodeDom.Compiler.CompilerParameters;
   $CompilerParameters.CompilerOptions = "/unsafe /platform:x86";
   $CompilerParameters.OutputAssembly = $WRAPPER_LIB_FILE;
   Add-Type -TypeDefinition $WrapperSourceCode -Language CSharp -CompilerParameters $CompilerParameters;

}

# Is this a Wow64 powershell host
Function Test-Wow64() {
    Return ((Test-Win32) -And (test-path env:\PROCESSOR_ARCHITEW6432))
}

# Is this a 64 bit process
Function Test-Win64() {
    Return ([IntPtr]::Size -Eq 8)
}

# Is this a 32 bit process
Function Test-Win32() {
    Return ([IntPtr]::Size -Eq 4)
}

Function Get-NetHASPData {
   Param (
      [Parameter(Mandatory = $true, ValueFromPipeline = $true)] 
      [String]$Command,
      [Switch]$SkipScanning,
      [Switch]$ReturnPlainText
   );
   # Interoperation to NetHASP stages:
   #    1. Set configuration (point to .ini file)
   #    2. Scan servers while STATUS not OK or Timeout not be reached
   #    3. Do one or several GET* command   

   # Init connect to NetHASP module?   
   $Result = "";
   if (-Not $SkipScanning) {
      # Processing stage 1
      Write-Verbose "$(Get-Date) Stage #1. Initializing NetHASP monitor session"
      $Result = ([HASP.Monitor]::doCmd("SET CONFIG,FILENAME=$HSMON_INI_FILE")).Trim();
      if ('OK' -Ne $Result) { 
         Write-Warning -Message "Error 'SET CONFIG' command: $Result"; 
         Return
      }
   
      # Processing stage 2
      Write-Verbose "$(Get-Date) Stage #2. Scan NetHASP servers"
      $Result = [HASP.Monitor]::doCmd("SCAN SERVERS");
      $ScanSeconds = 0;
      Do {
         # Wait a second before check process state
         Start-Sleep -seconds 1
         $ScanSeconds++; $Result = ([HASP.Monitor]::doCmd("STATUS")).Trim();
         #Write-Verbose "$(Get-Date) Status: $ret"
      } While (('OK' -ne $Result) -And ($ScanSeconds -Lt $HSMON_SCAN_TIMEOUT))

      # Scanning timeout :(
      If ($ScanSeconds -Eq $HSMON_SCAN_TIMEOUT) {
            Write-Warning -Message "'SCAN SERVERS' command error: timeout reached";
        }
    }

   # Processing stage 3
   Write-Verbose "$(Get-Date) Stage #3. Execute '$Command' command";
   $Result = ([HASP.Monitor]::doCmd($Command)).Trim();

   If ('EMPTY' -eq $Result) {
      Write-Warning -Message "No data recieved";
   } else {
      if ($ReturnPlainText) {
        # Return unparsed output 
        $Result;
      } else {
        # Parse output and push PSObjects to output
        # Remove double-quotes and processed lines that taking from splitted by CRLF NetHASP answer. 
        ForEach ($Line in ($Result -Replace "`"" -Split "`r`n" )) {
           If (!$Line) {Continue;}
           # For every non-empty line do additional splitting to Property & Value by ',' and add its to hashtable
           $Properties = @{};
           ForEach ($Item in ($Line -Split ",")) {
              $Property, $Value = $Item.Split('=');
              # "HS" subpart workaround
              if ($Null -Eq $Value) { $Value = "" }
              $Properties.$Property = $Value;
           } 
           # Return new PSObject with hashtable used as properties list
           New-Object PSObject -Property $Properties;
        }
      }
   }
}

<#
.SYNOPSIS  
    Return Sentinel/Aladdin HASP Network Monitor metrics value, make LLD-JSON for Zabbix

.DESCRIPTION
    Return Sentinel/Aladdin HASP Network Monitor metrics value, make LLD-JSON for Zabbix

.NOTES  
    Version: 1.2.1
    Name: Aladdin HASP Network Monitor Miner
    Author: zbx.sadman@gmail.com
    DateCreated: 18MAR2016
    Testing environment: Windows Server 2008R2 SP1, Powershell 2.0, Aladdin HASP Network Monitor DLL 2.5.0.0 (hsmon.dll)

    Due _hsmon.dll_ compiled to 32-bit systems, you need to provide 32-bit environment to run all code, that use that DLL. You must use **32-bit instance of PowerShell** to avoid runtime errors while used on 64-bit systems. Its may be placed here:_%WINDIR%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe.

.LINK  
    https://github.com/zbx-sadman

.PARAMETER Action
    What need to do with collection or its item:
        Discovery - Make Zabbix`'s LLD JSON;
        Get       - Get metric from collection`'s item;
        Count     - Count collection`'s items.
        DoCommand - Do NetHASP Monitor command that not require connection to server (HELP, VERSION). Command must be specified with -Key parameter

.PARAMETER ObjectType
    Define rule to make collection:
        Server - NetHASP server (detected with "GET SERVERS" command);
        Slot - NetHASP key slot ("GET SLOTS ...");
        Module - NetHASP module ("GET MODULES ...");
        Login - authorized connects to NetHASP server ("GET LOGINS ...").

.PARAMETER Key
    Define "path" to collection item`'s metric 

.PARAMETER ServerID
    Used to select NetHASP server from list. 
    ServerID can be numeric (real ID) or alphanumeric (server name)
    Server name must be taked from field "NAME" of the "GET SERVERS" command output ('stuffserver.contoso.com' or similar).

.PARAMETER ModuleID
    Used to additional objects selecting by Module Address

.PARAMETER SlotID
    Used to additional objects selecting by Slot

.PARAMETER LoginID
    Used to additional objects selecting by login Index

.PARAMETER ErrorCode
    What must be returned if any process error will be reached

.PARAMETER ConsoleCP
    Codepage of Windows console. Need to properly convert output to UTF-8

.PARAMETER Verbose
    Enable verbose messages

.EXAMPLE 
    Invoke-NetHasp -Action "DoCommand" -Key "VERSION"

    Description
    -----------  
    Get output of NetHASP Monitor VERSION command

.EXAMPLE 
    ... -Action "Discovery" -ObjectType "Server" 

    Description
    -----------  
    Make Zabbix`'s LLD JSON for NetHASP servers

.EXAMPLE 
    ... -Action "Get" -ObjectType "Slot" -Key "CURR" -ServerId "stuffserver.contoso.com" -SlotId "16" -ErrorCode "-127"

    Description
    -----------  
    Return number of used licenses on Slot #16 of stuffserver.contoso.com server. If processing error reached - return "-127"  

.EXAMPLE 
    ... -Action "Get" -ObjectType "Module" -Verbose

    Description
    -----------  
    Show formatted list of 'Module' object(s) metrics. Verbose messages is enabled. Console width is not changed.
#>
function Invoke-NetHasp
{
    Param (
       [Parameter(Mandatory = $True)] 
       [ValidateSet('DoCommand', 'Discovery', 'Get', 'Count')]
       [String]$Action,
       [Parameter(Mandatory = $False)]
       [ValidateSet('Server', 'Module', 'Slot', 'Login')]
       [Alias('Object')]
       [String]$ObjectType,
       [Parameter(Mandatory = $False)]
       [String]$Key,
       [Parameter(Mandatory = $False)]
       [String]$ServerId,
       [Parameter(Mandatory = $False)]
       [String]$ModuleId,
       [Parameter(Mandatory = $False)]
       [String]$SlotId,
       [Parameter(Mandatory = $False)]
       [String]$LoginId,
       [Parameter(Mandatory = $False)]
       [String]$ErrorCode = '-127',
       [Parameter(Mandatory = $False)]
       [String]$ConsoleCP,
       [Parameter(Mandatory = $False)]
       [String]$HSMON_LIB_PATH,
       [Parameter(Mandatory = $False)]
       [String]$HSMON_INI_FILE,
       [Parameter(Mandatory = $False)]
       [Int]$HSMON_SCAN_TIMEOUT = 30,
       [Parameter(Mandatory = $False)]
       [Switch]$JSON = $false
    );

    # Set default values from '1CHelper' module
    if ( -not $HSMON_LIB_PATH ) {
        $HSMON_LIB_PATH = (Get-NetHaspDirectoryPath).Replace('\','\\')
    }

    if ( -not $HSMON_INI_FILE ) {
        $HSMON_INI_FILE = (Get-NetHaspIniFilePath).Replace('\','\\')
    }

    #Set-StrictMode –Version Latest

    # Set US locale to properly formatting float numbers while converting to string
    #[System.Threading.Thread]::CurrentThread.CurrentCulture = "en-US";

    # Width of console to stop breaking JSON lines
    <#if ( -not $CONSOLE_WIDTH ) {
        Set-Variable -Name "CONSOLE_WIDTH" -Value 255 -Option Constant;
    }#>

    # Full paths to hsmon.dll and nethasp.ini
    if ( -not $HSMON_LIB_FILE ) {
        Set-Variable -Name "HSMON_LIB_FILE" -Value "$HSMON_LIB_PATH\\hsmon.dll" -Option Constant;
    }
    # Set-Variable -Name "HSMON_INI_FILE" -Value "$HSMON_LIB_PATH\\nethasp.ini" -Option Constant;

    # Full path to hsmon.dll wrapper, that compiled by this script
    if ( -not $WRAPPER_LIB_FILE ) {
        Set-Variable -Name "WRAPPER_LIB_FILE" -Value "$HSMON_LIB_PATH\\wraphsmon.dll" -Option Constant;
    }

    # Timeout in seconds for "SCAN SERVERS" connection stage
    Set-Variable -Name "HSMON_SCAN_TIMEOUT" -Value $HSMON_SCAN_TIMEOUT # -Option Constant;

    # Enumerate Objects. [int][NetHASPObjectType]::DumpType equal 0 due [int][NetHASPObjectType]::AnyNonexistItem equal 0 too
    Add-Type -TypeDefinition "public enum NetHASPObjectType { DumpType, Server, Module, Slot, Login }";

    Write-Verbose "$(Get-Date) Checking runtime environment...";

    # Script running into 32-bit environment?
    If ($False -Eq (Test-Wow64)) {
       Write-Warning "You must run this script with 32-bit instance of Powershell, due wrapper interopt with 32-bit Windows Library";
       Write-Warning "Try to use %WINDIR%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe `"&{Invoke-NetHasp [-Options ...]}`" [-OtherOptions ...]";
       Return;
    }

    Write-Verbose "$(Get-Date) Checking wrapper library for HASP Monitor availability...";
    If ($False -eq (Test-Path $WRAPPER_LIB_FILE)) {
       Write-Verbose "$(Get-Date) Wrapper library not found, try compile it";
       Write-Verbose "$(Get-Date) First wrapper library loading can get a few time. Please wait...";
       New-WrapperDLL;
       If ($False -Eq (Test-Path $WRAPPER_LIB_FILE)) {
        Write-Warning "Wrapper library not found after compilation. Something wrong";
        Return
       }
    } else {
      Write-Verbose "$(Get-Date) Loading wrapper library";
      Add-Type -Path $WRAPPER_LIB_FILE;
    }

    # Need to run one HASP command like HELP, VERSION ?
    If ('DoCommand' -Eq $Action) {
       If ($Key) {
           Write-Verbose "$(Get-Date) Just do command '$Key'";
           ([HASP.Monitor]::doCmd($Key)).Trim();
       } Else {
          Write-Warning -Message "No HASPMonitor command given with -Key option";
       }
       Return;
    }

    $Keys = $Key.Split(".");

    # Exit if object is not [NetHASPObjectType]
    #If (0 -Eq [int]($ObjectType -As [NetHASPObjectType])) { Exit-WithMessage -Message "Unknown object type: '$ObjectType'" -ErrorCode $ErrorCode; }

    Write-Verbose "$(Get-Date) Creating collection of specified object: '$ObjectType'";
    # Object must contain Servers data?
    if (($ObjectType -As [NetHASPObjectType]) -Ge [NetHASPObjectType]::Server) {
       Write-Verbose "$(Get-Date) Getting server list";
       $Servers = Get-NetHASPData -Command "GET SERVERS"; 
       if (-Not $Servers) { 
          Write-Warning -Message "No NetHASP servers found";
          Return;
       }

       Write-Verbose "$(Get-Date) Checking server ID";
       if ($ServerId) {
          # Is Server Name into $ServerId
          if (![RegEx]::IsMatch($ServerId,'^\d+$')) {
             # Taking real ID if true
             Write-Verbose "$(Get-Date) ID ($ServerId) was not numeric - probaly its hostname, try to find ID in servers list";
             $ServerId = (PropertyEqualOrAny -InputObject $Servers -Property NAME -Value $ServerId).ID;
             if (!$ServerId) {
                Write-Warning -Message "Server not found";
                Return
             }
             Write-Warning "$(Get-Date) Got real ID ($ServerId)";
          }
       }
       Write-Verbose "$(Get-Date) Filtering... (ID=$ServerId)";
       $Objects = $Servers = PropertyEqualOrAny -InputObject $Servers -Property ID -Value $ServerId
    }

    # Object must be processed with Servers data?
    if (($ObjectType -As [NetHASPObjectType]) -ge [NetHASPObjectType]::Module) {
       Write-Verbose "$(Get-Date) Getting modules list"; 
       $Modules = ForEach ($Server in $Servers) { 
          Get-NetHASPData -Command "GET MODULES,ID=$($Server.ID)" -SkipScanning; 
       }
       $Objects = $Modules = PropertyEqualOrAny -InputObject $Modules -Property MA -Value $ModuleId
    }

    # Object must be processed with Servers+Modules data?
    if (($ObjectType -As [NetHASPObjectType]) -ge [NetHASPObjectType]::Slot) {
       Write-Verbose "$(Get-Date) Getting slots list";
       $Slots = ForEach ($Module in $Modules) { 
          Get-NetHASPData -Command "GET SLOTS,ID=$($Module.ID),MA=$($Module.MA)" -SkipScanning; 
       }
       $Objects = $Slots = PropertyEqualOrAny -InputObject $Slots -Property SLOT -Value $SlotId
    }

    # Object must be processed with Servers+Modules+Slots data?
    If (($ObjectType -As [NetHASPObjectType]) -Ge [NetHASPObjectType]::Login) {
       Write-Verbose "$(Get-Date) Getting logins list";
       # LOGININFO ignore INDEX param and return list of Logins anyway
       $Logins = ForEach ($Slot In $Slots) { 
          Get-NetHASPData -Command "GET LOGINS,ID=$($Slot.ID),MA=$($Slot.MA),SLOT=$($Slot.SLOT)" -SkipScanning;
       }
       $Objects = $Logins | PropertyEqualOrAny -InputObject $Slots -Property INDEX -Value $LoginId
    }


    ForEach ($Object in $Objects) {   
      Add-Member -InputObject $Object -MemberType NoteProperty -Name "ServerName" -Value (PropertyEqualOrAny -InputObject $Servers -Property ID -Value $Object.ID).Name;
      Add-Member -InputObject $Object -MemberType AliasProperty -Name "ServerID" -Value ID
    }

    Write-Verbose "$(Get-Date) Collection created, begin processing its with action: '$Action'";
    switch ($Action) {
       'Discovery' {
          [Array]$ObjectProperties = @();
          Switch ($ObjectType) {
              'Server' {
                 $ObjectProperties = @("SERVERNAME", "SERVERID");
              }
              'Module' {
                 # MA - module address 
                 $ObjectProperties = @("SERVERNAME", "SERVERID", "MA", "MAX");
              }
              'Slot'   {
                 $ObjectProperties = @("SERVERNAME", "SERVERID", "MA", "SLOT", "MAX");
              }
              'Login' {
                 $ObjectProperties = @("SERVERNAME", "SERVERID", "MA", "SLOT", "INDEX", "NAME");
              }
           }
           if ( $JSON ) {
            Write-Verbose "$(Get-Date) Generating LLD JSON";
            $Result =  Get-NetHaspJSON <#Make-JSON#> -InputObject $Objects -ObjectProperties $ObjectProperties -Pretty;
           } else {
            $Result = $Objects
           }
       }
       'Get' {
          If ($Keys) { 
             Write-Verbose "$(Get-Date) Getting metric related to key: '$Key'";
             $Result = Format-ToZabbix -InputObject (Get-Metric -InputObject $Objects -Keys $Keys) -ErrorCode $ErrorCode;
          } Else { 
             Write-Verbose "$(Get-Date) Getting metric list due metric's Key not specified";
             $Result = <#Out-String -InputObject#> Get-Metric -InputObject $Objects -Keys $Keys;
          };
       }
       # Count selected objects
       'Count' { 
          Write-Verbose "$(Get-Date) Counting objects";  
          # if result not null, False or 0 - return .Count
          $Result = $(if ($Objects) { @($Objects).Count } else { 0 } ); 
       }
    }  

    # Convert string to UTF-8 if need (For Zabbix LLD-JSON with Cyrillic chars for example)
    <#If ($consoleCP) { 
       Write-Verbose "$(Get-Date) Converting output data to UTF-8";
       $Result = $Result | ConvertTo-Encoding -From $consoleCP -To UTF-8; 
    }#>

    # Break lines on console output fix - buffer format to 255 chars width lines 
    <#If (!$DefaultConsoleWidth) { 
       Write-Verbose "$(Get-Date) Changing console width to $CONSOLE_WIDTH";
       mode con cols=$CONSOLE_WIDTH; 
    }#>

    Write-Verbose "$(Get-Date) Finishing";
    $Result;
}

<#
.SYNOPSIS  
    Return USB (HASP) Device metrics value, count selected objects, make LLD-JSON for Zabbix

.DESCRIPTION
    Return USB (HASP) Device metrics value, count selected objects, make LLD-JSON for Zabbix

.NOTES  
    Version: 1.2.1
    Name: USB HASP Keys Miner
    Author: zbx.sadman@gmail.com
    DateCreated: 18MAR2016
    Testing environment: Windows Server 2008R2 SP1, USB/IP service, Powershell 2.0

.LINK  
    https://github.com/zbx-sadman

.PARAMETER Action
    What need to do with collection or its item:
        Discovery - Make Zabbix`'s LLD JSON;
        Get       - Get metric from collection item;
        Count     - Count collection items.

.PARAMETER ObjectType
    Define rule to make collection:
        USBController - "Physical" devices (USB Key)
        LogicalDevice - "Logical" devices (HASP Key)

.PARAMETER Key
    Define "path" to collection item`'s metric 

.PARAMETER PnPDeviceID
    Used to select only one item from collection

.PARAMETER ErrorCode
    What must be returned if any process error will be reached

.PARAMETER Verbose
    Enable verbose messages

.EXAMPLE 
    Invoke-UsbHasp -Action "Discovery" -ObjectType "USBController"

    Description
    -----------  
    Make Zabbix`'s LLD JSON for USB keys

.EXAMPLE 
    ... -Action "Count" -ObjectType "LogicalDevice"

    Description
    -----------  
    Return number of HASP keys

.EXAMPLE 
    ... -Action "Get" -ObjectType "USBController" -PnPDeviceID "USB\VID_0529&PID_0001\1&79F5D87&0&01" -ErrorCode "-127" -DefaultConsoleWidth -Verbose

    Description
    -----------  
    Show formatted list of 'USBController' object metrics selected by PnPId "USB\VID_0529&PID_0001\1&79F5D87&0&01". 
    Return "-127" when processing error caused. Verbose messages is enabled. 

    Note that PNPDeviceID is unique for USB Key, ID - is not.
#>
function Invoke-UsbHasp
{
    Param (
       [Parameter(Mandatory = $True)] 
       [ValidateSet('Discovery','Get','Count')]
       [String]$Action,
       [Parameter(Mandatory = $False)]
       [ValidateSet('LogicalDevice','USBController')]
       [Alias('Object')]
       [String]$ObjectType,
       [Parameter(Mandatory = $False)]
       [String]$Key,
       [Parameter(Mandatory = $False)]
       [String]$PnPDeviceID,
       [Parameter(Mandatory = $False)]
       [String]$ErrorCode = '-127',
       [Parameter(Mandatory = $False)]
       [Switch]$JSON = $false
    );

    # split key
    $Keys = $Key.Split(".");

    Write-Verbose "$(Get-Date) Taking Win32_USBControllerDevice collection with WMI"
    $Objects = Get-WmiObject -Class "Win32_USBControllerDevice";

    Write-Verbose "$(Get-Date) Creating collection of specified object: '$ObjectType'";
    Switch ($ObjectType) {
       'LogicalDevice' { 
          $PropertyToSelect = 'Dependent';    
       }
       'USBController' { 
          $PropertyToSelect = 'Antecedent';    
       }
    }

    # Need to take Unique items due Senintel used multiply logical devices linked to physical keys. 
    # As a result - double "physical" device items into 'Antecedent' branch
    #
    # When the -InputObject parameter is used to submit a collection of items, Sort-Object receives one object that represents the collection.
    # Because one object cannot be sorted, Sort-Object returns the entire collection unchanged.
    # To sort objects, pipe them to Sort-Object.
    # (C) PoSh manual
    $Objects = $( ForEach ($Object In $Objects) { 
                     PropertyEqualOrAny -InputObject ([Wmi]$Object.$PropertyToSelect) -Property PnPDeviceID -Value $PnPDeviceID
               }) | Sort-Object -Property PnPDeviceID -Unique;

    Write-Verbose "$(Get-Date) Processing collection with action: '$Action' ";
    Switch ($Action) {
       # Discovery given object, make json for zabbix
       'Discovery' {
          Write-Verbose "$(Get-Date) Generating LLD JSON";
          $ObjectProperties = @("NAME", "PNPDEVICEID");
          if ( $JSON ) {
            $Result = Make-JSON -InputObject $Objects -ObjectProperties $ObjectProperties -Pretty;
          } else {
            $Result = $Objects
          }
       }
       # Get metrics or metric list
       'Get' {
          If ($Keys) { 
             Write-Verbose "$(Get-Date) Getting metric related to key: '$Key'";
             $Result = PrepareTo-Zabbix -InputObject (Get-Metric -InputObject $Objects -Keys $Keys) -ErrorCode $ErrorCode;
          } Else { 
             Write-Verbose "$(Get-Date) Getting metric list due metric's Key not specified";
             $Result = $Objects;
          };
       }
       # Count selected objects
       'Count' { 
          Write-Verbose "$(Get-Date) Counting objects";  
          # if result not null, False or 0 - return .Count
          $Result = $( If ($Objects) { @($Objects).Count } Else { 0 } ); 
       }
    }

    Write-Verbose "$(Get-Date) Finishing";
    $Result;
}

# https://github.com/zbx-sadman
#endregion

Set-Alias -Name rac -Value Invoke-RAC -Description 'Клиент сервера администрирования 1С:Предприятие 8.3' -Scope Global
Set-Alias -Name ras -Value Invoke-RAS -Description "Сервер администрирования 1С:Предприятие 8.3" -Scope Global

Export-ModuleMember Remove-1CNotUsedObjects, Find-1CEstart, Find-1C8conn, Get-1ClusterData, Get-1CNetHaspIniStrings, Invoke-NetHasp, Invoke-UsbHasp, Remove-1CSession, Invoke-SqlQuery, Get-1CTechJournalData, Get-1CAPDEXinfo, Get-1CTechJournalLOGtable, Remove-1CTempDirs, Find-1CApplicationForExportImport, Get-1CHostData, Invoke-RAC, Get-1CAppDirs, Get-1CRegisteredApplicationClasses, Set-RACversion, New-RASservice, Set-RASversion, Invoke-RAS, Get-PerfCounters, Get-DiskSpdFromGitHub
