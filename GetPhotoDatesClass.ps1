#usage
#[PhotoMove]::("Source Directory","DestinationDirectory")
$photomove = [photomove]::new("$env:USERPROFILE\Desktop\PHONE")
$photomove = [photomove]::new("$env:USERPROFILE\DandiPhoto")

if ($PSVersionTable.PSVersion.Major -lt 5) {
  Write-Host "Requires PowerShell version 5 or higher."
  start "https://www.microsoft.com/en-us/download/confirmation.aspx?id=50395"
  Start-Sleep 120
  break;
}

class photomove{
  [string]$dateTimeRegex = '[0-9]{4}(0[1-9]|1[0-2])([1-2][0-9]|3[0-1])-(0[0-9]|1[0-9]|2[0-4])([0-6][0-9]){2}' #match 00000000-000000
  [bool]$rename = $true
  [bool]$createFolders = $true
  [string]$destinationPath = $env:USERPROFILE + "\Pictures"
  [string]$destinationFileName
  [string]$path = $env:USERPROFILE + "\Desktop\Phone"
  [string]$jpgDateTaken
  [string]$createDate
  [object]$extensionsToInclude = "*.png","*.gif","*.jpg","*.mov","*.3gp","*.avi","*.mp4","*.bmp","*.arw"

  PhotoMove () {
    $this.RenamePhotos()
  }
  PhotoMove ([string]$path) {
    $this.path = $path
    $this.RenamePhotos()
  }

  PhotoMove ([string]$path,[string]$destinationPath) {
    $this.path = $path
    $this.destinationPath = $destinationPath
    $this.RenamePhotos()
  }

  PhotoMove ([string]$path,[string]$destinationPath,[bool]$createFolders) {
    $this.path = $path
    $this.destinationPath = $destinationPath
    $this.createFolders = $createFolders
    $this.RenamePhotos()
  }

  #rename photos to date time
  [void] RenamePhotos () {

    $this.CheckPath($this.path)
    $files = Get-ChildItem -Recurse -Include $this.extensionsToInclude -Path $this.path

    foreach ($file in $files) {
      $dateTaken = $null
      $date = Get-Date $file.LastWriteTime -Format yyyyMMdd-HHmmss

      if ($file.FullName.ToLower().Contains(".jpg")) {
        $this.GetDateTakenJPG($file)

        if ($this.jpgDateTaken) {
          $this.MoveToDestinationPath($file,$this.jpgDateTaken)
        } else {
          $this.MoveToDestinationPath($file,$date)
        }
      }

      else {
        $this.MoveToDestinationPath($file,$date)
      }
    }
  }

  #move file
  [void] MoveToDestinationPath ($file,$date) {
    $destination = $null
    $targetFile = $null

    if ($this.createFolders = $true) {
      $year = $date.Substring(0,4)
      $month = $date.Substring(4,2)
      $destination = $this.destinationPath + "\$year\$month"

      if ((Test-Path $destination) -eq $false) {
        New-Item -ItemType Directory -Path $destination -Force
      }

      $targetFile = $destination + "\" + $date + "." + $file.FullName.Split(".")[1].ToLower()

    }

    if ((Test-Path $targetFile) -eq $false) {
      Write-Host "Rename with create date $file $targetFile" -ForegroundColor Yellow
      Move-Item -Path $file.FullName -Destination $targetFile
    }
    elseif (Test-Path $targetFile) {
      $this.AddDash($file,$targetFile)
    } else {
      Write-Host "No changes to $file"
    }
  }

  #if another file has the same name append _0001-9999
  [void] AddDash ($file,$targetFile) {
    $count = 1
    $targetFileNew = $null

    while (Test-Path $targetFile) {
      Write-Host $targetFile

      switch -regex ($count) {
        "\d{1}"
        {
          $targetFileNew = $targetFile.Replace(".","_000$count.")
        }
        "\d{2}"
        {
          $targetFileNew = $targetFile.Replace(".","_00$count.")
        }
        "\d{3}"
        {
          $targetFileNew = $targetFile.Replace(".","_0$count.")
        }
        "\d{4}"
        {
          $targetFileNew = $targetFile.Replace(".","_$count.")
        }
      }

      if ((Test-Path $targetFileNew) -eq $false) {
        $targetFile = $targetFileNew
      }

      $count++

    }

    if ($this.rename -eq $true) {
      Move-Item -Path $file.FullName -Destination $targetFile
      Write-Host "Add dash " $file.FullName $targetFileNew -ForegroundColor Green
    }
  }

  #get the date taken from the exif data
  [void] GetDateTakenJPG ($file) {
    $date = $null
    $fullpath = $file.FullName.ToLower().ToString()

    if ((Test-Path $fullpath) -and ($fullpath.Contains(".jpg"))) {
      [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms");
      $obj = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $fullpath

      try {
        $date = $obj.GetPropertyItem(36867).value[0..18]
      }
      catch {}

      if ($date -ne $null) {
        $arYear = [char]$date[0],[char]$date[1],[char]$date[2],[char]$date[3]
        $arMonth = [char]$date[5],[char]$date[6]
        $arDay = [char]$date[8],[char]$date[9]
        $arHour = [char]$date[11],[char]$date[12]
        $arMinute = [char]$date[14],[char]$date[15]
        $arSecond = [char]$date[17],[char]$date[18]
        $strYear = [string]::Join("",$arYear)
        $strMonth = [string]::Join("",$arMonth)
        $strDay = [string]::Join("",$arDay)
        $strHour = [string]::Join("",$arHour)
        $strMinute = [string]::Join("",$arMinute)
        $strSecond = [string]::Join("",$arSecond)
        $this.jpgDateTaken = $strYear + $strMonth + $strDay + "-" + $strHour + $strMinute + $strSecond
      } else {
        $this.jpgDateTaken = $null
      }

      $obj.Dispose()

    }
  }

  #try to avoid certain folders
  [void] CheckPath ($path) {
    if ($path -eq "c:\" -or $path.Contains("program files") -or $path.Contains("c:\windows") -or $path.Contains("programdata") -or $path.Contains(":") -eq $false) {
      Write-Output "Not allowed to run from this path $path"
      break;
    }
  }
}
