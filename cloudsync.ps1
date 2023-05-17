using assembly './WebDriver.dll'
using namespace OpenQA.Selenium
using namespace OpenQA.Selenium.Edge
using namespace OpenQA.Selenium.Support.UI
using namespace OpenQA.Selenium.Interactions
using namespace OpenQA.Selenium.Remote
using namespace System

function GetElementFromPath {
    [OutputType([WebElement])]
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$XPath,
        [int]$TimeSpanSeconds = 15,
        [int]$TimeSpanMilliseconds = 200
    )
    $TimeSpanWait = [TimeSpan]::FromSeconds($TimeSpanSeconds)
    $TimeSpanPoll = [TimeSpan]::FromMilliseconds($TimeSpanMilliseconds)

    $DriverWait = New-Object WebDriverWait($EdgeDriver, $TimeSpanWait)
    $DriverWait.PollingInterval = $TimeSpanPoll
    $DriverWait.IgnoreExceptionTypes([NoSuchElementException], [ElementNotVisibleException])

    $By = [By]::XPath($XPath)
    $Condition = {param([EdgeDriver]$x) 
        try {
            $Element = $x.FindElement($By)        
            if ($Element.Displayed) {
                Return $Element
            }
        } catch { return $null }
    }

    [WebElement]$Result = $null

    try {
        $Result = $DriverWait.Until[WebElement]($Condition)
    } catch { 
        Write-Host "$XPath Not Found"
        Return $null
    }

    Return $Result
}

function CreateDriver {
    [OutputType([EdgeDriver])]
    param (
        [string]$StartUrl
    )
    $EdgeService = [EdgeDriverService]::CreateDefaultService()
    $EdgeService.HideCommandPromptWindow = $true

    $EdgeOptions = New-Object EdgeOptions
    $EdgeOptions.AddArguments("headless")
    
    [EdgeDriver]$EdgeDriver = New-Object EdgeDriver($EdgeService, $EdgeOptions) 
    
    $Detector = New-Object LocalFileDetector
    $EdgeDriver.FileDetector = $Detector

    $EdgeDriver.Navigate().GoToURL($StartUrl)

    Return $EdgeDriver
}

function SendKeysTo {
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$XPath,
        [string]$Keys
    )

    $TextboxUser = GetElementFromPath $EdgeDriver -XPath $XPath
    $TextboxUser.SendKeys($Keys)
}
function SendPasswordAndSubmit {
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$Message,
        [string]$XPath,
        [ConsoleColor]$MessageColor
    )

    Write-Host -NoNewline -ForegroundColor $MessageColor $Message
    $Password = Read-Host -AsSecureString

    $TextboxPassword = GetElementFromPath $EdgeDriver -XPath $XPath
    $TextboxPassword.SendKeys((ConvertFrom-SecureString -SecureString $Password -AsPlainText))
    $Password.Clear()
    $TextboxPassword.Submit()
}

function Login {
    param (
        [EdgeDriver]$EdgeDriver
    )
    Write-Host -NoNewline -ForegroundColor Blue "Enter the username: "
    $Username = Read-Host

    SendKeysTo $EdgeDriver -XPath '//*[@id="username"]' -Keys $Username

    SendPasswordAndSubmit $EdgeDriver -Message "Enter the password for ${Username}: " -XPath '//*[@id="password"]' -MessageColor Red

    SendPasswordAndSubmit $EdgeDriver -Message "Enter the mailbox password: " -XPath '//*[@id="mailboxPassword"]' -MessageColor White
}

function UploadFolder {
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$FolderPath #'C:\Users\Downloads'
    )
    $ByFolder = [By]::XPath('//input[@type="file"][@webkitdirectory="true"]') 
    $InputElementFolder = $EdgeDriver.FindElement($ByFolder) 
    $InputElementFolder.SendKeys($FolderPath)

    $Js = "document.querySelectorAll('input[type=\'file\'][webkitdirectory=\'true\']').forEach(element => element.value='');"
    $EdgeDriver.ExecuteScript($Js)
}

function GoToFolder {
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$FolderName
    )

    $ActionCustom = New-Object Actions($EdgeDriver)

    $ByiCloud = [By]::XPath("//td[.//span[contains(text(), 'Folder - $FolderName')]]")
    $FolderiCloud = $EdgeDriver.FindElement($ByiCloud)

    $ActionCustom.DoubleClick($FolderiCloud).Perform()
}

function GoToRoot {
    param (
        [EdgeDriver]$EdgeDriver
    )

    $ByRoot = [By]::XPath('//button[@title="My files"]')
    $ButtonRoot = $EdgeDriver.FindElement($ByRoot)
    $ButtonRoot.Click()
}

function UploadFile {
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$FilePath
    )

    $BySingle = [By]::XPath('//input[@type="file"][@multiple]') 
    $InputElementSingle = $EdgeDriver.FindElement($BySingle) 
    $InputElementSingle.SendKeys($FilePath)

    $Js = "document.querySelectorAll('input[type=\'file\'][multiple]').forEach(element => element.value='');"
    $EdgeDriver.ExecuteScript($Js)
}

function DeleteOldest {
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$FileName,
        [int]$ArchivesToKeep
    )

    try {        
        $ByZip = [By]::XPath("//tr[.//span[contains(text(), 'File - application/zip - ${FileName}')]]") 
        $ZipRows = $EdgeDriver.FindElements($ByZip) 
        if ($ZipRows.Count -gt $ArchivesToKeep) {
            $ToDelete = $ZipRows | Sort-Object -Property ComputedAccessibleLabel | Select-Object -First 1
            $ByCheckbox = [By]::XPath(".//td[.//input[@type='checkbox']]") 
            $DeleteCheckbox = $ToDelete.FindElement($ByCheckbox)
            $DeleteCheckbox.Click()

            $ByDelete = [By]::XPath("//button[@data-testid='toolbar-trash']") 
            $DeletButton = $EdgeDriver.FindElement($ByDelete)
            $DeletButton.Click()
            $FileDeletedName = (($ToDelete.Text  -split '\n')[0] -split ' - ')[2]
            Write-Host "File $FileDeletedName deleted."
        }

        Write-Host "$($ZipRows.Count) archives founds. Not deleting."
    } catch { Write-Host "Nothing to delete."}
}

function CompressFolder {
    [OutputType([string])]
    param (
        [string]$FolderName
    )
    $OutputFileSuffix = (Get-Date).ToString("yyyy-MM-ddTHH.mm.ss") + ".zip"
    $OutputFilePath = "$env:TEMP\$FolderName" + $OutputFileSuffix
    Compress-Archive -CompressionLevel "Fastest" -Path "$env:USERPROFILE\$FolderName\*" -DestinationPath $OutputFilePath

    Return $OutputFilePath
}

function DeleteFiles {
    param (
        [string[]]$FilePaths
    )
     ForEach ($FilePath in $FilePaths) {
        If (Test-Path $FilePath) {
            Remove-Item -Path $FilePath -Force
        }
    }
}

function CloseDriver {
    param (
        [EdgeDriver]$EdgeDriver
    )
    $EdgeDriver.Close()
    $EdgeDriver.Quit()
}

function WaitToUpload {
    param (
        [EdgeDriver]$EdgeDriver,
        [int]$MinutesToWait,
        [int]$SecondsToSleep
    )
    $IsComplete = $false
    $ByStatus = [By]::XPath("//div[contains(@class, 'transfers-manager-list-item')][.//progress]")
    $ItemStatuses = $EdgeDriver.FindElements($ByStatus) 

    $EllapsedSeconds = 0

    while(-Not $IsComplete)
    {
        if ($EllapsedSeconds -gt $MinutesToWait * 60) {
            break
        }

        $AllDone = $true

        foreach ($ItemStatus in $ItemStatuses)
        {
            $AllDone = $ItemStatus.Text.Contains("Uploaded") -And $AllDone
        }

        $IsComplete =  $AllDone
        Start-Sleep -Seconds $SecondsToSleep
        $EllapsedSeconds = $EllapsedSeconds + $SecondsToSleep
    }

    $ByClose = [By]::XPath("//button[@data-testid='drive-transfers-manager:close']") 
    $CloseButton = $EdgeDriver.FindElements($ByClose) 
    $CloseButton.Click()
}

$iCloudDrive = "iCloudDrive"
$OneDrive = "OneDrive"
$ProtonDriveUrl = 'https://drive.proton.me'
$MinutesToWait = 60
$SecondsToSleep = 5
$ArchivesToKeep = 5

$iCloudZipPath = CompressFolder $iCloudDrive
$OneDriveZipPath = CompressFolder $OneDrive

$EdgeDriver = CreateDriver $ProtonDriveUrl

Login $EdgeDriver

GoToFolder $EdgeDriver -FolderName $iCloudDrive
UploadFile $EdgeDriver -FilePath $iCloudZipPath
DeleteOldest $EdgeDriver -FileName $iCloudDrive -ArchivesToKeep $ArchivesToKeep

GoToRoot $EdgeDriver

GoToFolder $EdgeDriver -FolderName $OneDrive
UploadFile $EdgeDriver -FilePath $OneDriveZipPath
DeleteOldest $EdgeDriver -FileName $iCloudDrive -ArchivesToKeep $ArchivesToKeep

WaitToUpload $EdgeDriver -MinutesToWait $MinutesToWait -SecondsToSleep $SecondsToSleep

DeleteFiles $iCloudZipPath, $OneDriveZipPath

CloseDriver $EdgeDriver