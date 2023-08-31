<#
.SYNOPSIS
    Скрипт для получения информации о казначейских счетах и иных реквизитах для перевода денежных средств в бюджетную систему Российской Федерации.
.DESCRIPTION
    Сначала с использованием API необходимо постранично загрузить данные о казначейских счетах. Далее форматируем данные, удаляя ненужные свойства, и результатом станет файл в формате csv,
    
    который выгружается по указанному пути. Если указать параметр -TRSAPath, то дополнительно из данных о казначейских счетах будет создан новый объект trsa с нужными свойствами, 
    
    будет загружен zip архив с Банковскими Идентификационными Кодами, разархивирован, и на его основе будет создано новое свойство BankName в объекте trsa. Результатом станет файл в формате csv, который 
    
    выгружается по указанному пути.  
.LINK
    https://github.com/Solbreen/Kaznacheistvo/blob/main/KAZNA/KazFastmm.ps1
#>

[CmdletBinding()]
param (
    #   Путь файла с данными казначейских счетов.
    #   Введите полный путь и имя файла(в формате .csv)
    [Parameter(Mandatory)]
    [string]$KSPath,

    #   Введите количество одновременно выполняемых потоков. По умолчанию значение 6.
    [Parameter()]
    [int32]$JobLimit = 6,

    #   Путь файла с данными Банковских Идентификационных Кодов.
    #   Введите полный путь и имя файла(в формате .csv)
    [Parameter()]
    [string]$TRSAPath,

    #   Введите URI сетевого прокси-сервера, если вы используете прокси-сервер.
    [Parameter()]
    [uri]$Proxy
)

#region preferences
$data = @()
$urlBIK = "https://www.cbr.ru/s/newbik"
$SB = {
    param([array]$par, $pro)
    $urlKS = "http://budget.gov.ru/epbs/registry/7710568760-KRKS/data?pageSize=20&pageNum=%page%&filteractual=true&filterstatus=ACTIVE&filter_start_accountnum=40102810&filter_not_stateks=KS3"
    $data_json = @()
    $requestparam = @{
        'uri' = $urlKS
    }
    if ($PSBoundParameters.ContainsKey('pro')) {
        $requestparam.add('proxy', $pro)
        $requestparam.add('ProxyUseDefaultCredentials', $true)
    }
    foreach ($jtem in $par) {
        $requestparam.uri = $urlks.Replace("%page%", "$jtem")
        $data_json += Invoke-RestMethod @requestparam
    }
    $data_json
}
#endregion

$starttime = get-date
Write-Verbose "Начато: $(get-date)" 

#region create paths
#  Проверяем наличие указанных путей, куда будут загружаться файлы с данными.
#  Если таких путей нет, то программа сама их создаст.
if($PSBoundParameters.ContainsKey('TRSAPath')){
    [array]$DopPath = $KSPath, $TRSAPath
} 
else { 
    [array]$DopPath = $KSPath
}
foreach ($path in $DopPath) {
    $rootpath = split-path -path $Path -parent
    if ( -not (test-path $rootpath)) {
        try {
            New-Item -path $rootpath -itemType Directory -ErrorAction Stop | Out-Null
            Write-Verbose "Путь '$rootpath' создан"
        }
        catch {
            Write-Warning "Не удалось создать каталог '$rootpath' для выгрузки файла."
            throw "$($_.Exception.Message)"
        }
    }
}
#endregion

#region Запрос первой страницы
$firstpageparam = @(1)
if ($PSBoundParameters.ContainsKey('proxy')) {
    $firstpageparam += $Proxy
}
try {
    $firstpage = invoke-command $sb -ArgumentList $firstpageparam -ErrorAction Stop
}
catch {
    Write-Warning "Не удалось скачать первую страницу"
    throw "$($_.Exception.Message)"
}
$data += $firstpage.data
$pagecount = $firstpage.pagecount
write-verbose "Общее число страниц для скачивания: $pagecount"
#endregion

#region Создание массива с массивом чисел
$PageRange = [math]::round(($pagecount * 4) / 100)
$mod = $pagecount % $PageRange
$numofit = ($pagecount - $mod) / $PageRange
$Nmas = @(1..($numofit + 1))
$stp = 1

for ($i = 1; $i -le $numofit; ++$i) {
    $enp = $i * $PageRange
    if ($i -eq 1) {
        $smas = @(2..($enp))    
    }
    else {
        $smas = @($stp..($enp))   
    }
    $Nmas[$i - 1] = $smas
    $stp = 1 + $enp
}
$nmas[$numofit] = @($stp..$pagecount)
write-verbose "Количество блоков страниц для скачивания: $($nmas.count)" 
#endregion

#region Запускаем скачивание и выгрузку данных
Write-Verbose "Очищаем место от старых задач"
get-job | remove-job

Write-Verbose "Запускаем задачи с лимитом одновременных выполнений: $JobLimit"
foreach ($item in $Nmas) {
    while ((get-job -state Running).count -ge $JobLimit) {
        Write-Progress -Activity "Этап 1/2. Выгружаем данные с сайта" -Status "Блоков: $((get-job).count) из $($nmas.count).   $([math]::round(($((get-job).count) / $($nmas.count) * 100), 1))%" -PercentComplete ($((get-job).count) / $($nmas.count) * 100)
        start-sleep -Seconds 1 
    }
    Write-Verbose "Запускаем задачу для обработки страниц с $($item[0]) по $($item[-1]) "
    if ($PSBoundParameters.ContainsKey('proxy')) {
        Start-Job -ScriptBlock $SB -ArgumentList $item, $proxy | Out-Null
    }
    else {
        Start-Job -ScriptBlock $SB -ArgumentList (, $item) | Out-Null
    }
}    

while ((get-job -state Completed).count -ne $nmas.count) {
    Write-Progress -Activity "Этап 2/2. Ожидаем завершения работ" -Status "Завершено: $((get-job -state Completed).count) из $($nmas.count).   $([math]::round(($((get-job -state Completed).count) / $($nmas.count) * 100), 1))%" -PercentComplete ($((get-job -state Completed).count) / $($nmas.count) * 100)
} 
write-verbose "Получаем результаты задач"
$data += (receive-job -job (get-job) -Keep).data

if ($data.count -ne $firstpage.recordcount) {
    Write-Warning "Обработанное количество записей ($($data.count)) не совпадает с количеством записей на ресурсе ($($firstpage.recordcount))."
    throw "Не совпадает число обработанных записей с записями на ресурсе"
}

Write-Verbose "Загрузка данных осуществлена за время : $((get-date).Add(-$starttime).TimeOfDay.tostring())"
try {
    $data | select-object -property ksnumber, 
    @{ label = "clientname"; expr = { $_.clientname.replace("`n", "") } }, 
    accountnum, 
    tofkbik, 
    @{ label = "dateopen"; expr = { get-date -Date $_.dateopen -Format "dd.MM.yyyy" } }, 
    opentofkcode, 
    signls | Export-Csv -Path $KSPath -Encoding utf8 -NoTypeInformation -ErrorAction stop
    Write-Verbose "Выгрузили файл с казначейскими счертами в файл '$KSPath'"
}
catch {
    Write-Warning "Не удаётся выгрузить файл с казначейскими счетами."
    throw "$($_.Exception.Message)"
}
#endregion

#region TRSA
if ($PSBoundParameters.ContainsKey('TRSAPath')) {
    #  Выделяем единые казначейские счета в отдельный справочник
    $tempkatalog = "$PSScriptroot\$(get-random -minimum 100000000 -maximum 100000000000000000)"
    $keeptrsa = @( 'accountnum', 'tofkbik', 'opentofkcode' )
    $trsa = @()
    $trsa = ($data | Select-object -Unique -Property $keeptrsa)
    Write-Verbose "Количество записей в справочнике trsa $($trsa.count)"

    # Создаём временный каталог
    try {
        new-item -path "$tempkatalog" -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Verbose "Создан каталог для скачивания архива с БИК '$tempcatalog'"
    }
    catch {
        Write-Warning "Не удалось создать временный каталог '$tempcatalog'"
        throw "$($_.Exception.Message)"
    }

    $bikparam = @{
        'uri' = $urlBIK
    }
    if ($PSBoundParameters.ContainsKey('proxy')) {
        $bikparam.add('proxy', $Proxy)
        $bikparam.add('ProxyUseDefaultCredentials', $true) 
    }


    #  Качаем zip файл с данным БИК, распаковываем из него xml документ с данными 
    $bikfile = "$tempkatalog\bik.zip"
    Write-Verbose "Запускаем скачивание БИК архива в файл '$bikfile'"
    try {
        Invoke-WebRequest @bikparam -OutFile $bikfile 
    }
    catch {
        Write-Warning "БИК архив скачать не удалось"
        Remove-Item -Path $tempkatalog
        throw "$($_.Exception.Message)"
    }
    Expand-Archive -path $bikfile  -DestinationPath $tempkatalog

    # Получаем данные из xml файла, создаём пространство имен для работы с xml файлом
    $xmlfile = "$tempkatalog\*.xml"
    [xml]$xml = get-content -path $xmlfile
    Write-Verbose "Подгрузили данные из скачанного xml файла '$xmlfile'"
    $ns = 'urn:cbr-ru:ed:v2.0'
    write-Verbose "Создаём объект с типом Xml.XmlNamespaceManager для добавления namespace '$ns'"
    $Namespace = New-Object -TypeName "Xml.XmlNamespaceManager" -ArgumentList $XML.NameTable
    $Namespace.AddNamespace('ed', $ns)

    foreach ($item in $trsa) {
        $bic = $item.tofkbik
        $a = $XML.SelectNodes("//ed:BICDirectoryEntry[@BIC='$bic']", $Namespace) # Выбираем из xml файла часть кода c БИК территориального органа Федерального казначейства
        $hasAccounts = $false 
        foreach ($entry in $a) {
            $NameCBR = ''
            $el_info = $entry.selectsinglenode("ed:ParticipantInfo", $namespace) # Для удобства работы назначаем переменную огбъектом со свойствами participantinfo текущего БИК
            $NameP = $el_info.namep
            $tnp = $el_info.tnp
            $nnp = $el_info.nnp
            [array]$el_acc = $entry.SelectNodes("ed:Accounts", $Namespace) # Добавляем в массив все элементы свойства accounts из части кода текущего БИК
            $hasAccounts = $el_acc.count -gt 0 # проверка на существование элементов в accounts. Если вызвать переменную, то вернется булево значение
            if ($hasAccounts) {
                $AccountCBRBIC = $el_acc[0].AccountCBRBIC
                foreach ($i in $XML.SelectNodes("//ed:BICDirectoryEntry[@BIC='$AccountCBRBIC']", $Namespace)) {
                    # ищем cbrbik по xml файлу
                    $el_info = $i.selectsinglenode("ed:ParticipantInfo", $namespace)
                    $NameCBR = $el_info.namep
                }
            }
        }
        # Добавляем в элемент trsa свойство BankName и присваивается значение, зависящее от значения $hasAccounts
        if ($hasAccounts) {
            add-member -InputObject $item -Name "BankName" -MemberType NoteProperty -Value "$NameCBR//$NameP, $tnp $nnp"
        }
        else {
            add-member -InputObject $item -Name "BankName" -MemberType NoteProperty -Value "$NameCBR, $tnp $nnp"
        }
    }

    try {
        $trsa | Export-Csv -Path $TRSAPath -Encoding utf8 -NoTypeInformation -ErrorAction Stop
        Write-Verbose "Выгрузили справочник trsa в файл '$TRSAPath'"
    }
    catch {
        Write-Warning "Не удалось выгрузить справочник trsa в файл '$TRSAPath'"
        throw "$($_.Exception.Message)"
    }
    finally {
        Write-Verbose "Удаляем временный каталог '$tempkatalog'"
        Remove-Item -Path $tempkatalog -Recurse
    }
}
#endregion

Write-Verbose "Закончено : $(get-date))" 