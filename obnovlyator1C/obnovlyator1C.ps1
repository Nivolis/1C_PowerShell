# 
# 
# 
#====================================================================================

# Задаем кодовую страницу
chcp.com 1251

# !!! для отладки
Clear-Host


# запись в лог файл
function Write-Log($logFile)
{ 
  Process { 
    Write-Verbose -Message $_
    if ($logFile)
    {
      (Get-Date -Format yyy_MM_dd__HH_mm_ss) + ' ' + $_ | Out-File -FilePath $logFile -Append
    }
  }
}


# разбираем вывод от RAC в объекты
function RacOutToObject($rac_out)
{
  $objectList = @()  
  $object     = New-Object -TypeName PSObject      

  FOREACH ($line in $rac_out) 
  {
    #Write-Host "raw: _ $line _"

    if (([string]::IsNullOrEmpty($line))) 
    {
      $objectList += $object
      $object     = New-Object -TypeName PSObject 
    }

    # Remove the whitespace at the beginning on the line
    $line = $line -replace '^\s+', ''
   
    $keyvalue = $line -split ':'
	
    $key     = $keyvalue[0] -replace '^\s+', ''
    $value   = $keyvalue[1] -replace '^\s+', ''

    $key	 = $key.trim() -replace '-', '_'
    $value = $value.trim()

    if (-not ([string]::IsNullOrEmpty($key))) 
    {
      $object | Add-Member -Type NoteProperty -Name $key -Value $value
    }
  }

  return $objectList
}

# блокируем базу на вход через RAC
function RACBaseLock($PathRac, $ServerNameRAC, $BaseName, $UserName, $UserPass, $UcCode, $logFile)
{
  $cluster_uuid = (RacOutToObject (& $PathRac $ServerNameRAC cluster list)).cluster
  $infobases = RacOutToObject (& FPathRac $ServerNameRAC infobase --cluster=$cluster_uuid summary list)
  
  # ищем ID базы
  FOREACH ($infobase in $infobases)
  {
    if ($infobase.name -eq $BaseName)
    {
      $infobase_uuid = $infobase.infobase
    }
  }
  
  # блокируем базу на вход
  Write-Host -Object "Блокируем базу $BaseName в кластере $ServerNameRAC на вход" | Write-Log -logFile $logFile
  & $PathRac $ServerNameRAC infobase update --cluster=$cluster_uuid --infobase=$infobase_uuid --infobase-user=$UserName --infobase-pwd=$UserPass --sessions-deny=on --denied-message="" --denied-from="" --permission-code=$UcCode
  Write-Host -Object "Блокируем базу $BaseName в кластере $ServerNameRAC на вход" | Write-Log -logFile $logFile
}


# снимаем блокировку базы на вход через RAC
function RACBaseUnLock($PathRac, $ServerNameRAC, $BaseName, $UserName, $UserPass, $logFile)
{
  $cluster_uuid = (RacOutToObject(& $PathRac $ServerNameRAC cluster list)).cluster
  $infobases = RacOutToObject (& $PathRac $ServerNameRAC infobase --cluster=$cluster_uuid summary list)
  
  # ищем ID базы
  FOREACH ($infobase in $infobases)
  {
    if ($infobase.name -eq $BaseName) 
    {
      $infobase_uuid = $infobase.infobase
    }
  }
  
  # снимаем блокировку с базы на вход
  & $PathRac $ServerNameRAC infobase update --cluster=$cluster_uuid --infobase=$infobase_uuid --infobase-user=$userName --infobase-pwd=$UserPass --sessions-deny=off --denied-message="" --denied-from="" --permission-code=""
}


# выключаем регламентные задания в базе через RAC
function RACBaseJobsLock($PathRac, $ServerNameRAC, $BaseName, $UserName, $UserPass, $logFile)
{
  $cluster_uuid = (RacOutToObject(& $PathRac $ServerNameRAC cluster list)).cluster
  $infobases = RacOutToObject (& $PathRac $ServerNameRAC infobase --cluster=$cluster_uuid summary list)
  
  # ищем ID базы
  FOREACH ($infobase in $infobases)
  {
    if ($infobase.name -eq $BaseName)
    {
      $infobase_uuid = $infobase.infobase
    }
  }
  
  # выключаем регламентные задания в базе
  & $PathRac $ServerNameRAC infobase update --cluster=$cluster_uuid --infobase=$infobase_uuid --infobase-user=$userName --infobase-pwd=$UserPass --scheduled-jobs-deny=on
}


function RACBaseJobsUnLock($PathRac, $ServerNameRAC, $BaseName, $UserName, $UserPass, $logFile)
{
  $cluster_uuid = (RacOutToObject(& $PathRac $ServerNameRAC cluster list)).cluster
  $infobases = RacOutToObject (& $PathRac $ServerNameRAC infobase --cluster=$cluster_uuid summary list)
  
  # ищем ID базы
  FOREACH ($infobase in $infobases)
  {
    if ($infobase.name -eq $BaseName)
    {
      $infobase_uuid = $infobase.infobase
    }
  }
  
  # выключаем регламентные задания в базе
  & $PathRac $ServerNameRAC infobase update --cluster=$cluster_uuid --infobase=$infobase_uuid --infobase-user=$userName --infobase-pwd=$UserPass --scheduled-jobs-deny=off
}


# разрываем сеансы всех пользователей в базе через RAC
# можно еще через, /C"ЗавершитьРаботуПользователей" / /C"РазрешитьРаботуПользователей" . но не во всех конфигурациях есть данная процедура
function RACTerminateAllUsers($PathRac, $ServerNameRAC, $BaseName, $logFile)
{
  $cluster_uuid = (RacOutToObject(& $PathRac $ServerNameRAC cluster list)).cluster
  $infobases = RacOutToObject (& $PathRac $ServerNameRAC infobase --cluster=$cluster_uuid summary list)
  
  # ищем ID базы
  FOREACH ($infobase in $infobases)
  {
    if ($infobase.name -eq $BaseName)
    {
      $infobase_uuid = $infobase.infobase
    }
  } 
 
  # получаем все сессии в указанной базе
  $sessions = RacOutToObject(& $PathRac $ServerNameRAC session list --cluster=$cluster_uuid --infobase=$infobase_uuid)
  
  FOREACH ($session in $sessions)
    {
      $session_uuid = $session.session
      $sessionsUsr = $session.user_name
    
      # пишем какие сеансы закрыли
      Write-Host -Object "Закрываем: $sessionsUsr сеанс в базе: $BaseName" | Write-Log $logFile
      # здесь вывод всей информации о сеансе
      $session | Format-List | Write-Log $logFile
      & $PathRac $ServerNameRAC session terminate --cluster=$cluster_uuid --session=$session_uuid | Write-Log $logFile
    }
}


# выполняет бекап базы в DT
function BackUpBaseToDT($Path1C, $ServerName1C, $BaseName, $UserName, $UserPass, $UcCode, $logFile)
{
  $NOWDATETIME = Get-Date -Format yyy_MM_dd__HH_mm_ss
  
  #Сформируем имя файла резервной копии
  $BAKFN = $PathBackup + $BaseName + "_" + $NOWDATETIME + ".dt"
  
  #Бекап базы в .dt  
  $arglist = "DESIGNER /S$ServerName1C\$BaseName /N$UserName /P$UserPass /UC$UcCode /DumpIB $BAKFN /Visible /Out $logFile -NoTruncate"
  Start-Process -FilePath $Path1C -ArgumentList $arglist -Wait
}


# выполняем обновление конфигурации
function UpdateCf($Path1C, $ServerName1C, $BaseName, $UserName, $UserPass, $UcCode, $PathCfu, $logFile)
{
  # /LoadCfg для конф снятых с поддержки,
  # /UpdateCfg обновление конфигурации для конфигурации надящейся на поддержке  
  $arglist = "DESIGNER /S$ServerName1C\$BaseName /N$UserName /P$UserPass /UC $UcCode /UpdateCfg $PathCfu /Visible /Out $logFile -NoTruncate"
  Start-Process -FilePath $Path1C -ArgumentList $arglist -Wait
  
  &$Path1C 'DESIGNER' '/S' $ServerName1C'\'$BaseName '/N' $UserName '/P' $UserPass '/UpdateCfg' $PathCfu '/Visible' '/Out' $logFile '-NoTruncate' '-Server' '/UC' $UcCode
}


# выполняем обновление базы
function UpdateDB($Path1C, $ServerName1С, $BaseName, $UserName, $UserPass, $logFile)
{
  # обновление базы с запуском 1С в режиме предприятия после завершения обновления
  $arglist = "DESIGNER /S$ServerName1С\$BaseName /N$UserName /P$UserPass /UpdateDBCfg -WarningsAsErrors -Server /Visible /Out $logFile -NoTruncate"
  Start-Process -FilePath $Path1C -ArgumentList $arglist -Wait
}

function PSVersionMinimum
{
  if ($PSVersionTable.PSVersion.Major -clt 5) 
  {
    Write-Host -Object 'версия Poreshell устарела, обновите до 5.1' | Write-Log($logFile)
    Exit
  }
}




#===================================================================================================================
# START
#===================================================================================================================
# файл с параметрами в переменную
$CSVFile = Get-Content -Path 'C:\Users\ss_vershinin\Desktop\config_obn.csv' | ConvertFrom-Csv -Delimiter ';'

# построчно перебираем параметры из файла
foreach ($parmetrInCsv in $CSVFile) 
{
  # считываем параметры в переменные
  $ServerNameRAC = $parmetrInCsv.ServerNameRAC
  $ServerName1C = $parmetrInCsv.ServerName1C
  $BaseName = $parmetrInCsv.BaseName
  $UserName = $parmetrInCsv.UserName
  $UserPass = $parmetrInCsv.UserPass
  $UcCode = $parmetrInCsv.UcCode
  $PathCfu = $parmetrInCsv.PathCfu
  $PathBackup = $parmetrInCsv.PathBackup
  $PathRac = $parmetrInCsv.PathRac
  $logFile = $parmetrInCsv.LogFile
  $Path1C = $parmetrInCsv.Path1C
  
  #проверка версии Powershell, использовать не ниже 5
  PSVersionMinimum | Write-Log($logFile)
  
  #Блокировка сеансов на вход rac
  RACBaseLock $PathRac $ServerNameRAC $BaseName $UserName $UserPass $UcCode $logFile
    
  #Отключаем регламентные задания rac
  RACBaseJobsLock $PathRac $ServerNameRAC $BaseName $UserName $UserPass $logFile
  
  # разрываем сеансы всех пользователей в базе через RAC
  RACTerminateAllUsers $PathRac $ServerNameRAC $BaseName $logFile
  
  #Резервная копия базы
  BackUpBaseToDT $Path1C $ServerName1C $BaseName $UserName $UserPass $UcCode $logFile
  
  #Обновление конфигурации
  
  
  #Обновление базы  
  
  
  # запускаем 1С в режиме предприятия, для обновления базы и ждем закрытия
  $arglist = "ENTERPRISE /S$ServerName1C\$BaseName /N$UserName /P$UserPass /UC $UcCode"
  Start-Process -FilePath $Path1C -ArgumentList $arglist -Wait
  
  #Включаем регламентные задания rac
  RACBaseJobsUnLock $PathRac $ServerNameRAC $BaseName $UserName $UserPass $logFile
  
  #Снимаем блокировку сеансов на вход rac
  RACBaseUnLock $PathRac $ServerNameRAC $BaseName $UserName $UserPass $logFile
  
  # !!! для отладки
  Write-Host Здесь делаем паузу, для просмотра результатов работы
  Write-Host для продолжения любая клавиша.
  Write-Host для выхода Ctrl-C
  Pause
}
