function Remove-NotUsedObjects
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

function Find-1CEstart
<#
.Synopsis
   Поиск стартера 1С
.DESCRIPTION
   Поиск исполняемого файла 1cestart.exe
.EXAMPLE
   Find-1CEstart
.OUTPUTS
   NULL или строку с полным путём к исполняемому файлу
#>
{
    Param(
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
             Write-Warning "Не найдены ключи в: $($string.leaf)\\$($string.path)"
         }

         $defaultValue = $regkey.GetValue("").ToString()

         $index = $defaultValue.IndexOf("1cestart.exe")

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

    if ( -not $pathToStarter -and $ComputerName -ne $env:COMPUTERNAME ) {

        $pathToStarter = Invoke-Command -ComputerName $ComputerName -ScriptBlock { 
                            if ( Test-Path "${env:ProgramFiles(x86)}\1cv8\common\1cestart.exe" ) {
                                "${env:ProgramFiles(x86)}\1cv8\common\1cestart.exe" 
                            } elseif ( Test-Path "${env:ProgramFiles(x86)}\1cv82\common\1cestart.exe" ) {
                                "${env:ProgramFiles(x86)}\1cv82\common\1cestart.exe"
                            } else { $null } 
                         } -ErrorAction Continue

    } elseif ( -not $pathToStarter ) {

        $pathToStarter = if ( Test-Path "${env:ProgramFiles(x86)}\1cv8\common\1cestart.exe" ) {
                                "${env:ProgramFiles(x86)}\1cv8\common\1cestart.exe" 
                            } elseif ( Test-Path "${env:ProgramFiles(x86)}\1cv82\common\1cestart.exe" ) {
                                "${env:ProgramFiles(x86)}\1cv82\common\1cestart.exe"
                            } else { $null }
                              
    }

    $pathToStarter
}

function Find-1C8conn
<#
.Synopsis
   Поиск строк подключения 1С
.DESCRIPTION
   Поиск строк подключения
.EXAMPLE
   Find-1C8conn
.OUTPUTS
   массив найденных строк поключения 1С
#>
{
    [OutputType([Object[]])]
    Param(
        # Использовать общие файлы
        [switch]$UseCommonFiles = $true,
        [string[]]$UseFilesFromDirectories
    )

    # TODO http://yellow-erp.com/page/guides/adm/service-files-description-and-location/
    # TODO http://yellow-erp.com/page/guides/adm/service-files-description-and-location/

    $list = @()

    # TODO 

    $list
    
}

function Get-LicensesStatistic
<#
.Synopsis
   Собирает статистику использования лицензий по серверам (не ниже 8.3)
.DESCRIPTION
   
.EXAMPLE
   Get-LicensesStatistic
.EXAMPLE
   Get-LicensesStatistic 'srv-01','srv-02'
.OUTPUTS
   Таблицу статискики
#>
{
    [OutputType([Object])]
    Param(
        # Использовать список компьютеров из файла nethasp.ini
        [switch]$UseNetHasp=$false,
        # Адреса компьютеров для проверки лицензий
        [string[]]$Hosts
    )

    $hostsToQuery = @()

    if ( $UseNetHasp ) {
        $netHaspParams = Get-NetHaspIniStrings
        $hostsToQuery += $netHaspParams.NH_SERVER_ADDR
        $hostsToQuery += $netHaspParams.NH_SERVER_NAME
    }

    if ( $Hosts ) {
        $hostsToQuery += $Hosts
    }

    $connector = New-Object -ComObject "v83.COMConnector"
    
    if ( $connector ) {
        # TODO
    }

}

function Get-NetHaspIniStrings
<#
.Synopsis
   Находит значения параметров в файле nethasp.ini
.DESCRIPTION
   
.EXAMPLE
   Get-NetHaspIniStrings
.OUTPUTS
   Структура параметров
#>
{
    
    $struct = @{}
    
    $pathToStarter = Find-1CEstart
    $pathToFile = $pathToStarter.Replace("common\1cestart.exe", "conf\nethasp.ini")

    if ( $pathToStarter ) {
        
        $content = Get-Content -Encoding UTF8 -LiteralPath $pathToFile
        $strings = $content | ? { $_ -match "^\w" }
        $strings | % { $keyValue = $_.Split('='); $key = $keyValue[0].Replace(" ",""); $value = $keyValue[1].Replace(" ",""); $value = $value.Replace(';',''); $struct[$key] = $value.Split(',') }

    }

    $struct

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
        [string]$ComputerName=''
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
             Write-Warning "Не найдены ключи в: HKLM:\\$UninstallKey"
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

function Get-NetHaspDirectoryPath
<#
.Synopsis
   Возвращает путь к каталогу с библиотекой hsmon.dll
.DESCRIPTION
   
.EXAMPLE
   Get-NetHaspDirectoryPath
.OUTPUTS
   Путь к каталогу с библиотекой hsmon.dll
#>
{  
    (Get-Module 1CHelper).Path.Replace("\1CHelper.psm1", "")
}

function Get-NetHaspIniFilePath
<#
.Synopsis
   Возвращает путь к файлу nethasp.ini
.DESCRIPTION
   
.EXAMPLE
   Get-NetHaspIniFilePath
.OUTPUTS
   Путь к к файлу nethasp.ini
#>
{  
    $pathToStarter = Find-1CEstart
    $pathToStarter.Replace("common\1cestart.exe", "conf\nethasp.ini")
}

function Invoke-UsbHasp
<#
.Synopsis
   Вызывает на выполнение сценарий usbhasp.ps1 в текущем окне

.DESCRIPTION
   
.EXAMPLE
   Invoke-UsbHasp -Verbose
   
.EXAMPLE
   Invoke-UsbHasp -Discovery -Verbose

#>
{
    [CmdletBinding()]
    Param (
       [Parameter(Mandatory = $False)] 
       [ValidateSet('Discovery','Get','Count')]
       [String]$Action = 'Discovery',
       [Parameter(Mandatory = $False)]
       [ValidateSet('LogicalDevice','USBController')]
       [Alias('Object')]
       [String]$ObjectType = 'USBController',
       [Parameter(Mandatory = $False)]
       [String]$Key,
       [Parameter(Mandatory = $False)]
       [String]$PnPDeviceID,
       [Parameter(Mandatory = $False)]
       [String]$ErrorCode = '-127',
       [Parameter(Mandatory = $False)]
       [String]$ConsoleCP,
       [Parameter(Mandatory = $False)]
       [Switch]$DefaultConsoleWidth
    )

    $HSMON_LIB_PATH = (Get-NetHaspDirectoryPath).Replace('\','\\')
    $HSMON_INI_FILE = (Get-NetHaspIniFilePath).Replace('\','\\')
    
    . "$(Get-NetHaspDirectoryPath)\usbhasp.ps1" -Action:$Action -ObjectType:$ObjectType -Key:$Key  `
                                                -ErrorCode:$ErrorCode -ConsoleCP:$ConloleCP `
                                                -DefaultConsoleWidth:$DefaultConsoleWidth
}

<# BEGIN
https://github.com/zbx-sadman
#>

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

function Invoke-NetHasp
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
        Discovery - Make Zabbix's LLD JSON;
        Get       - Get metric from collection's item;
        Count     - Count collection's items.
        DoCommand - Do NetHASP Monitor command that not require connection to server (HELP, VERSION). Command must be specified with -Key parameter

.PARAMETER ObjectType
    Define rule to make collection:
        Server - NetHASP server (detected with "GET SERVERS" command);
        Slot - NetHASP key slot ("GET SLOTS ...");
        Module - NetHASP module ("GET MODULES ...");
        Login - authorized connects to NetHASP server ("GET LOGINS ...").

.PARAMETER Key
    Define "path" to collection item's metric 

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
    Make Zabbix's LLD JSON for NetHASP servers

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
{
    Param (
       [Parameter(Mandatory = $False)] 
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
       [Switch]$NoJSON = $true
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
       $Objects = $Logins = PropertyEqualOrAny -InputObject $Slots -Property INDEX -Value $LoginId
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
           if ( -not $NoJSON ) {
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

<# END
https://github.com/zbx-sadman
#>
