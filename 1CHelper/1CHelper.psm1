function Delete-NotUsedObjects
<#
.Synopsis
   Удаление неиспользуемых объектов конфигурации
.DESCRIPTION
   Удаление элементов конфигурации с синонимом "(не используется)"
.EXAMPLE
   PS C:\> $modules = Delete-NoUsedTypes E:\TEMP\ExportingConfiguration
   PS C:\> $gr = $modules | group File, Object | select -First 1
   PS C:\> ise ($gr.Group.File | select -First 1) # открываем модуль в новой вкладке ISE
   # альтернатива 'start notepad $gr.Group.File[0]'
   PS C:\> $gr.Group | select Object, Type, Line, Position -Unique | sort Line, Position | fl # Смотрим что корректировать
   PS C:\>  $modules = $modules | ? File -NE ($gr.Group.File | select -First 1) # удаление обработанного файла из списка объектов
   # альтернатива '$modules = $modules | ? File -NE $psise.CurrentFile.FullPath'
   # и все сначала с команды '$gr = $modules | group File, Object | select -First 1'
.INPUTS
   Пусть к файлам выгрузки конфигурации
.OUTPUTS
   Массив объектов с описанием файлов модулей и позиций, содержащих упоминания удаляемых объектов
#>
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
        66..90 | % { $chars += [char]$_ }
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
            $files = ls -LiteralPath $pathToFiles -Filter *.xml -File #| ? { $_.Name.ToString().Split('.').Count -eq 3 }
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
                        | measure).Count -gt 0) {
                    # Добавляем имя файла в массив удаляемых
                    $fileRefs += $item.FullName
                    # Находим производные типы
                    $tmp = $xml | Select-Xml -XPath ("//root:$pref/root:InternalInfo/readable:GeneratedType/@name") -Namespace $hashTable
                    $tmp | % { $typeRefs += $_.ToString() }
                    # Находим подчиненные объекты (<ИмяТипаПлатформы>.<ИмяТипаМетаданных>.*) и добавляем к удаляемым файлам
                    ls -LiteralPath $pathToFiles -Filter "$($m[0]).$($m[1]).*" -File | ? { $_.Name.ToString().Split('.').Count -gt 3 } | % { $fileRefs += $_.FullName }
                    $addAll = $true
                } elseif(-not $thisIsShort) {
                    
                }
                # Поиск аттрибутов
                if ($addAll) {
                    # Поиск аттрибутов
                    $xml.SelectNodes("//root:MetaDataObject/root:$pref/root:ChildObjects/root:Attribute/root:Properties/root:Name", $nsmgr) | % { $childRefs += "$pref.$name.Attribute.$($_.'#text')" } 
                    # Поиск форм
                    $xml.SelectNodes("//root:MetaDataObject/root:$pref/root:ChildObjects/root:Form", $nsmgr) | % { $childRefs += "$pref.$name.Form.$($_.'#text')" }
                    # Поиск команд
                    $xml.SelectNodes("//root:MetaDataObject/root:$pref/root:ChildObjects/root:Command/root:Properties/root:Name", $nsmgr) | % { $childRefs += "$pref.$name.Command.$($_.'#text')" }
                    # Поиск макетов
                    $xml.SelectNodes("//root:MetaDataObject/root:$pref/root:ChildObjects/root:Template", $nsmgr) | % { "$pref.$name.Template.$($_.'#text')" }
                    # Поиск ресурсов информациооного регистра
                    if ($pref -eq 'InformationRegister') {
                        $xml.SelectNodes("//root:MetaDataObject/root:$pref/root:ChildObjects/root:Resource/root:Properties/root:Name", $nsmgr) | % { "$pref.$name.Resource.$($_.'#text')" }
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
            $fileRefs | ? { $_ -notlike '' } | % {rm $_ -Verbose}
            # Выбираем оставшиеся для поиска неиспользуемых ссылок на типы и атрибутов
            Write-Progress -Activity "Поиск файлов *.xml" -Completed -Status "Подготовка"
            $filesToUpdate = ls -LiteralPath $pathToFiles -Filter *.xml -File
            # Удаляем пустой элемент (Создан при вызове конструктора типа)
            Write-Progress -Activity "Обработка ссылок для поиска" -Completed -Status "Подготовка"
            $typeRefs = $typeRefs | ? { $_ -notlike '' } | select -Unique
            $childRefs = $childRefs | ? { $_ -notlike '' } | select -Unique
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
                        $chars | % { $max = [Math]::Max($max, $tpref.LastIndexOf($_)) }
                        if ($max -eq -1) {
                            Write-Error "Неверный тип для поиска" -ErrorId C2 -Targetobject $tref -Category ParserError
                            continue
                        } else {
                            $type = if($max -eq 0) { $tref.Split('.')[0] } else { $tpref.Substring(0, $max) }
                            try {
                                $xml.SelectNodes("//root:MetaDataObject/root:Configuration/root:ChildObjects/root:$type[text()='$($tref.Split('.')[1])']/.", $nsmgr) `
                                    | % { $_.ParentNode.RemoveChild($_) | Out-Null }
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
                        $chars | % { $max = [Math]::Max($max, $tpref.LastIndexOf($_)) }
                        if ($max -eq -1) {
                            Write-Error "Неверный тип для поиска" -ErrorId S2 -Targetobject $tref -Category ParserError
                            continue
                        } else {
                            $type = if($max -eq 0) { $tref.Split('.')[0] } else { $tpref.Substring(0, $max) }
                            try {
                                $xml.SelectNodes("//root:MetaDataObject/root:Subsystem/root:Properties/root:Content/item:Item[text()='$($type+'.'+$tref.Split('.')[1])']/.", $nsmgr) `
                                    | % { $_.ParentNode.RemoveChild($_) | Out-Null }
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
                        $chars | % { $max = [Math]::Max($max, $tpref.LastIndexOf($_)) }
                        if ($max -eq -1) {
                            Write-Error "Неверный тип для поиска" -ErrorId S2 -Targetobject $tref -Category ParserError
                            continue
                        } else {
                            $type = if($max -eq 0) { $tref.Split('.')[0] } else { $tpref.Substring(0, $max) }
                            try {
                                $xml.SelectNodes("//exch:$name/exch:Item/exch:Metadata/[text()='$($type+'.'+$tref.Split('.')[1])']/.", $nsmgr) `
                                    | % { $_.ParentNode.RemoveChild($_) | Out-Null }
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
                    $typeRefs | % { $xml.SelectNodes("//*/core:Type[contains(text(), '$_')]/.", $nsmgr) } | % { $_.ParentNode.RemoveChild($_) | Out-Null }
                    Write-Verbose "Поиск неиспользуемых атрибутов"
                    $xml.SelectNodes("//*/core:content[contains(translate(text(),$($dict.replace),$($dict.with)),'(не используется)')]/../../../..", $nsmgr) `
                        | % { $_.ParentNode.RemoveChild($_) | Out-Null }
                }
                if (Test-Path -LiteralPath $item.FullName) {
                    $xml.Save($item.FullName)
                }
                $i++
            }
            # Обработка модулей объектов
            Write-Progress -Activity "Поиск файлов модулей (*.txt)" -Completed
            $txtFiles = ls -LiteralPath $pathToFiles -Filter *.txt -File
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

function Find-1CApplicationForExportImport
<#
.Synopsis
   Поиск максимальной версии приложения
.DESCRIPTION
   Поиск максимальной версии приложения (не ниже 8.3)
.EXAMPLE
   Find-1CApplicationForExportImport
.OUTPUTS
   NULL или строку с путем установки приложения
#>
{
    Param(
        # Имя компьютера для поиска версии
        [string]$computerName=''
    )
    $installationPath = $null

    $pvs = 0

    $UninstallPathes = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")
   
    ForEach($UninstallKey in $UninstallPathes) {
        
         Try {
             $reg=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computerName)
         } Catch {
             $_
             Continue
         }
 
         $regkey = $reg.OpenSubKey($UninstallKey)

         If(-not $regkey) {
             Write-Error "Не найдены ключи в: HKLM:\\$UninstallKey"
         }

         $subkeys=$regkey.GetSubKeyNames()
        
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
                 $installationPath = $tmpPath
             }
 
         }
         $reg.Close() 
     }
     $installationPath
}