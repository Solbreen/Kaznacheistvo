<#
.SYNOPSIS
    A short one-line action-based description, e.g. 'Tests if a function is valid'
.DESCRIPTION
    A longer description of the function, its purpose, common use cases, etc.
.NOTES
    Information or caveats about the function e.g. 'This function is not supported in Linux'
.LINK
    Specify a URI to a help page, this will show when Get-Help -Online is used.
.EXAMPLE
    Test-MyTestFunction -Verbose
    Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines
#>

#region preferences
$urlKS = "http://budget.gov.ru/epbs/registry/7710568760-KRKS/data?pageSize=20&pageNum=%page%&filteractual=true&filterstatus=ACTIVE&filter_start_accountnum=40102810&filter_not_stateks=KS3"
$urlBIK = "https://www.cbr.ru/s/newbik"
$pagenum = 1
$data = [System.Collections.ArrayList]::new()
$proxy = @{
    ProxyUseDefaultCredentials = $true
    proxy                      = 'Http://172.23.17.3:8080'
}
$KSPath = "c:\test\KSnew.csv"
$TRSAPath = "c:\test\TRSAnew.csv"
$tempkatalog = "$PSScriptroot\$(get-random -minimum 100000000 -maximum 100000000000000000)"
$CurrentVP = $VerbosePreference
$VerbosePreference = 'Continue'
#endregion



$starttime = get-date
write-host "Начато:"(get-date) -ForegroundColor Magenta

#region create paths
#  Проверяем наличие указанных путей, куда будут загружаться файлы с данными.
#  Если таких путей нет, то программа сама их создаст.
foreach ($path in $KSPath, $TRSAPath) {
    if ( -not (test-path (split-path -path $Path -parent))) {
        try {
            New-Item -path (split-path -path $Path -parent) -itemType Directory -ErrorAction Stop | Out-Null
            Write-host "Путь $(split-path -path $Path -parent) создан"
        }
        catch {
            throw "$($_.Exception.Message)"
        }
    }
}
#endregion


#region KS_Load
#  Этот блок кода отвечает за загрузку справочника казначейских счетов
do {
    try {
        $data_json = Invoke-RestMethod -Uri ($urlKS.Replace("%page%", "$pagenum")) -ErrorAction Stop @proxy -Verbose:$false
    }
    catch {
        throw "$($_.Exception.Message)"
    }
    
    if ( $pagenum -eq 1) {
        $recordscounttotal = $data_json.recordcount
        $pagecount = $data_json.pagecount
    }
    
    foreach($item in $data_json.data){
        [void]$data.add($item)
    }
    Write-Progress -Activity "Выгружаем данные с сайта" -Status "Страниц: $pagenum из $pagecount.   $([math]::round(($pagenum / $pagecount * 100), 1))%" -PercentComplete ($pagenum / $pagecount * 100)
    $pagenum = $pagenum + 1
} until ( $data.count -eq $recordscounttotal ) 
Write-host "Загрузка данных осуществлена за время :"(get-date).Add(-$starttime).TimeOfDay.tostring() -ForegroundColor Green
Write-host "recordscounttotal = " $recordscounttotal
#endregion

#region trsaformat
#  Выделяем единые казначейские счета в отдельный справочник
$keeptrsa = @( 'accountnum', 'tofkbik', 'opentofkcode' )
$trsa = @()
$trsa = ($data | Select-object -Unique -Property $keeptrsa)
#endregion

#region trsaformat with BIK
# Создаём временный каталог
try {
    new-item -path "$tempkatalog" -ItemType Directory -ErrorAction Stop | Out-Null
}
catch {
    throw "$($_.Exception.Message)"
}

#  Качаем zip файл с данным БИК, распаковываем из него xml документ с данными 
$bikfile = "$tempkatalog\bik.zip"
try {
    Invoke-WebRequest -Uri $urlBIK -ErrorAction Stop @proxy -OutFile $bikfile 
}
catch {
    throw "$($_.Exception.Message)"
}
Expand-Archive -path $bikfile  -DestinationPath $tempkatalog

# Получаем данные из xml файла, создаём пространство имен для работы с xml файлом
Write-Verbose "Подгружаем данные из скачанного xml файла"
[xml]$xml = get-content -path "$tempkatalog\*.xml"
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
            foreach ($i in $XML.SelectNodes("//ed:BICDirectoryEntry[@BIC='$AccountCBRBIC']", $Namespace)) { # ищем cbrbik по xml файлу
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
#endregion

#region end
#  Конец. Просто выгружаем файлы
try {
    $data | select-object -property ksnumber, 
    @{ label = "clientname"; expr = { $_.clientname.replace("`n", "") } }, 
    accountnum, 
    tofkbik, 
    @{ label = "dateopen"; expr = { get-date -Date $_.dateopen -Format "dd.MM.yyyy" } }, 
    opentofkcode, 
    signls | Export-Csv -Path $KSPath -Encoding utf8 -NoTypeInformation -ErrorAction stop
}
catch {
    throw "$($_.Exception.Message)" #под выражение или subexpression
}
try {
    $trsa | Export-Csv -Path $TRSAPath -Encoding utf8 -NoTypeInformation -ErrorAction Stop
}
catch {
    throw "$($_.Exception.Message)"
}
#endregion

Remove-Item -Path $tempkatalog -Recurse
$VerbosePreference = $CurrentVP 
Write-host "Закончено :"((get-date)) -ForegroundColor Magenta