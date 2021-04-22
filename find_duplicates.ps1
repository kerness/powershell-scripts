<#      Script allows to ccheck for duplicates in directory and monitor for new and modified files.
#       In input parameters you need to specified a path to directory to check
# 
#       It uses PSSQLite module created by RamblingCookieMonster - https://github.com/RamblingCookieMonster/PSSQLite
#       It is mandatory to have SQLite installed on Your machine!
#       PowerShell version: 7.1
#       Created by: Maciej Bąk, 400666 geoinf - 17.04.2021
#>

# input parameters
param(
  [string]$Inputpath = 'C:\Users\Public\Pictures' # default value
)




# This function was created to count hashes using Start-Job but due to unexpected behaviour with PSSQLite module I decided to don't use it
# function Count-Hashes {
#   param (
#       $pathToCount
#   )
#       # start new job for counting hashes
#   $Job = Start-Job -ScriptBlock {
#       Get-ChildItem -Path $args[0] -Recurse -File  | Select-Object DirectoryName, Name, FullName, LastWriteTime, Length, @{n = 'FileHash'; e = { (Get-FileHash -LiteralPath $_.FullName).Hash } } 
#   } -ArgumentList $pathToCount
#   # wait for it
#   Wait-Job  $Job
#   # get it
#   $res = Receive-Job $Job
#   # and remove it
#   Remove-Job $Job
#   $res = $res | Select-Object DirectoryName, Name, FullName, LastWriteTime, Length, FileHash
#   return $res
# }

function Get-AllDuplicates {
  # get duplicates and uniqueDuplicates
  # get list of all duplicated hashes
  param (
    $db,
    $tabName
  )
  $Q_allDuplicates ='
    SELECT * FROM "{0}" WHERE FileHash IN
      (SELECT FileHash FROM "{0}" GROUP BY FileHash HAVING COUNT(*) > 1)
  ' -f $tabName

  $allDuplicates_res = Invoke-SqliteQuery -DataSource $db -Query $Q_allDuplicates
  return $allDuplicates_res
}

function Get-uniqueDuplicates {
  # get list of only unique hashes values which are duplicated
  param (
    $db,
    $tabName
  )
    
  $Q_uniqueDuplicates ='
    SELECT FileHash, COUNT(FileHash)
    FROM "{0}"
    GROUP BY FileHash
    HAVING COUNT(FileHash)>1
  ' -f $tabName
  $UniqueDuplicates_res = Invoke-SqliteQuery -DataSource $db -Query $Q_uniqueDuplicates
  return $UniqueDuplicates_res
}
function Install-ModuleSmart {
  # check if module is installed or imported and install or load it if not
  param (
    [string]$module
  )
  if ($null -eq (Get-InstalledModule | Where-Object {$_.Name -eq $module})) { # if TRUE - module is NOT installed  (Get-InstalledModule returns all modules installed on the machine)
    Install-Module PSSQLite
    Import-Module PSSQLite
  }
  if ($null -eq (Get-Module | Where-Object {$_.Name -eq $module})) { # if TRUE - module IS installed but is NOT imported (Get-Module returns modules in current sesion)
    Import-Module PSSQLite
  }
  
  $currentModule = Get-InstalledModule | Where-Object {$_.Name -eq $module} | Select-Object Version, Name, Description
  Write-Host "Module nammed $($currentModule.Name) is installed correctly." -ForegroundColor green
  Write-Host "Version: $($currentModule.Version) - $($currentModule.Name) - $($currentModule.Description)"
}
function Write-HashesToHost {
  <#
    Prints parent hash and childs of that parent.
  #>
  param (
    $parents,
    $all
  )
  foreach ($parent in $parents) {
    $print = $all | Where-Object { $_.FileHash -eq $parent.FileHash }
    Write-Host "----------------------------------------------------" -ForegroundColor Black
    Write-Host "Plik o hashu $($parent.FileHash) znaleziono w następujących lokalizacjach: " -ForegroundColor Blue
      
    for ($i = 0; $i -lt $print.Count; $i++) {
      $string = "    - {0}" -f $print[$i].DirectoryName + $print[$i].Name
      $string
    }
  }
}
function Add-File {
  <#
    Check if file exists - if not create it if yes do nothing.
  #>
  param (
    [string]$path
  )

  $chcek = Test-Path -Path $path -PathType Leaf
  if (-not($chcek)) { # if file does not exist jus create it
    try {
      $newFile = New-Item -ItemType File -Path $path -Force
      return $newFile
    }
    catch {
      throw $_.Exception.Message
    }
  } else { # if exists do NOTHING
    return $path
  }
}

# install module to interact with SQLite database
Install-ModuleSmart -module "PSSQLite"
$DataBase = Add-File -path ".\find_duplicates_db.SQLite"
$path = $Inputpath

# name for TMP table
$tmpTable = $path + 'TMP'

# tables names list from database 
$Q_tables = '
  SELECT name FROM sqlite_master WHERE TYPE = "table" AND name NOT LIKE "sqlite_%"
'
$tablesNames = Invoke-SqliteQuery -Query $Q_tables -DataSource $DataBase


# Main script if-else statement - check if table for specified path exists (check if that path was used before to count hashes).

if ($tablesNames.Name -contains $path) { # there IS such path

  # get files
  $files = Get-ChildItem -Path $path -Recurse -File 
  $files = $files | Select-Object DirectoryName, Name, FullName, LastWriteTime, Length, @{n = 'FileHash'; e = { 0 } } # dodac creationTime

  # create temporary table
  $Q_createTableTMP = '
    CREATE TABLE IF NOT EXISTS "{0}" (
      id integer PRIMARY KEY,
      FileHash char(64) NOT NULL,
      DirectoryName text NOT NULL,
      LastWriteTime datetime NOT NULL,
      Name varchar(256) NOT NULL,
      FullName text NOT NULL,
      Length integer NOT NULL)
  ' -f $tmpTable
  Invoke-SqliteQuery -Query $Q_createTableTMP -DataSource $DataBase
  # push new files data to table
  $files = $files | Out-DataTable
  Invoke-SQLiteBulkCopy -DataTable $files -DataSource $DataBase -Table $tmpTable -NotifyAfter 5 -Force # path is also a table name
 
  # get all files that are new or which have been modified since the last checking - save it to #toCalc variable
  $Q_chooseFiles = '
    SELECT base.id, tmp.FileHash, tmp.DirectoryName, tmp.LastWriteTime, tmp.Name, tmp.FullName, tmp.Length
    FROM "{0}" tmp
    LEFT JOIN "{1}" base ON base.FullName = tmp.FullName
    WHERE base.name IS NULL OR base.LastWriteTime != tmp.LastWriteTime;
  ' -f $tmpTable, $path
  $toCalc = Invoke-SqliteQuery -Query $Q_chooseFiles -DataSource $DataBase
  Write-Host "Those are new or modified files" -ForegroundColor Green
  $toCalc | Select-Object FullName| Format-Table

  # calculate hashes for them
  $newHashes = $toCalc | Select-Object id, DirectoryName, Name, FullName, LastWriteTime, Length, @{n = 'FileHash'; e = { (Get-FileHash -LiteralPath $_.FullName).Hash } } # dodac creationTime
  Write-Host "Those are new hashes to push to database" -ForegroundColor Green
  $newHashes | Select-Object @{n = 'File path'; e = { $_.FullName }}, @{n = 'File Hash'; e = { $_.FileHash }} | Format-Table

  # now, we need to update our database - need to push $newHashes to '$path' named table (without tmp)
  # if file was modified id is NOT NULL and represents that specific file - so it will be easy - just update by ID
  foreach ($item in $newHashes) {
    if ($null -ne $item.Id) { # id is NOT NULL so the item.Id represents that file in '$path' named table
      Write-Host "One record has been updated."
      [DateTime]$date = $item.LastWriteTime
      $dateString = $date.ToString("yyyy-MM-dd HH:mm:ss") # make sure that the date format is correct (H is for 24h system)
      # query to update records
      $Q_notNull = '
        UPDATE "{0}" 
        SET FileHash = "{1}", LastWriteTime = "{2}", Length = {3}
        WHERE id = {4};
      '-f $path, $item.FileHash, $dateString, $item.Length, $item.id
      Invoke-SqliteQuery -Query $Q_notNull -DataSource $DataBase
    }else {  # when id is null is also easy because there is nothing special to do - just add files do '$path' named table 
      Write-Host "One record has been added."
      [DateTime]$date = $item.LastWriteTime
      $dateString = $date.ToString("yyyy-MM-dd HH:mm:ss") # make sure that the date format is correct (H is for 24h system)
      # query to add records
      $Q_null = '
        INSERT INTO "{0}" (FileHash, DirectoryName, LastWriteTime, Name, FullName, Length)
        VALUES ("{1}", "{2}", "{3}", "{4}", "{5}", {6});
      ' -f $path, $item.FileHash, $item.DirectoryName, $dateString, $item.Name, $item.FullName, $item.Length 
      Invoke-SqliteQuery -Query $Q_null -DataSource $DataBase
    }
  }

  # drop that tmp table
  $Q_dropTMP = 'DROP TABLE "{0}";' -f $tmpTable
  Invoke-SqliteQuery -Query $Q_dropTMP  -DataSource $DataBase

  # get all duplicates form data base (hash and how many)
  $allDuplicates_res = Get-AllDuplicates -db $DataBase -tabName $path
  # get only hashes which are unique in all duplicates list
  $uniqueDuplicates_res = Get-uniqueDuplicates -db $DataBase -tabName $path
  #print only if allDuplicates_res is not empty!
  if ($null -eq $allDuplicates_res) {
    Write-Host "Lucky man! You don't have any duplicates!" -ForegroundColor Green
  }
  else {
    Write-HashesToHost -parents $uniqueDuplicates_res -all $allDuplicates_res
  }

} else { # if there is NO such path

  Write-Host "New table will be created" -ForegroundColor Green
  # get files and count hashes
  $files = Get-ChildItem -Path $path -Recurse -File 
  $hashes = $files | Select-Object DirectoryName, Name, FullName, LastWriteTime, Length, @{n = 'FileHash'; e = { (Get-FileHash -LiteralPath $_.FullName).Hash } } # dodac creationTime

  #$hashes = Count-Hashes $path

  
  # table creation query. Create it only if it not exists
  $Q_createTable = '
  CREATE TABLE IF NOT EXISTS "{0}" (
    id integer PRIMARY KEY AUTOINCREMENT,
    FileHash char(64) NOT NULL,
    DirectoryName text NOT NULL,
    LastWriteTime datetime NOT NULL,
    Name varchar(256) NOT NULL,
    FullName text NOT NULL,
    Length integer NOT NULL
  )
  ' -f $path
  Invoke-SqliteQuery -Query $Q_createTable -DataSource $DataBase

  # insert hashes to database
  $hashes = $hashes | Out-DataTable
  Invoke-SQLiteBulkCopy -DataTable $hashes -DataSource $DataBase -Table $path -NotifyAfter 5 -Force # path is also a table name
  # get all duplicates form data base (which hash and how many)
  $allDuplicates_res = Get-AllDuplicates -db $DataBase -tabName $path
  # get only hashes which are unique in all duplicates list
  $uniqueDuplicates_res = Get-uniqueDuplicates -db $DataBase -tabName $path
  #print only if allDuplicates_res is not empty!
  if ($null -eq $allDuplicates_res) {
    Write-Host "Lucky man! You don't have any duplicates!" -ForegroundColor Green
  }
  else {
    Write-HashesToHost -parents $uniqueDuplicates_res -all $allDuplicates_res
  }
}
