using assembly './WebDriver.dll'
using namespace OpenQA.Selenium
using namespace OpenQA.Selenium.Edge
using namespace OpenQA.Selenium.Support.UI
using namespace OpenQA.Selenium.Interactions
using namespace OpenQA.Selenium.Remote
using namespace System

function GetElementFromPath {
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$XPath,
        [bool]$ReturnArray = $false,
        [bool]$IsHidden = $false,
        [WebElement]$ElementToCheck = $null,
        [int]$TimeSpanSeconds = 20,
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
            if ($null -eq $ElementToCheck) {
                $ElementToCheck = $x
            }

            if ($ReturnArray) {
                $Elements = $ElementToCheck.FindElements($By)        
                if ($Elements.Displayed -Or $IsHidden) {
                    Return $Elements
                }
            } else {
                $Element = $ElementToCheck.FindElement($By)        
                if ($Element.Displayed -Or $IsHidden ) {
                    Return $Element
                }
            }            
        } catch { return $null }
    }

    $Result = $null

    try {
        if ($ReturnArray) {
            $Result = $DriverWait.Until[WebElement[]]($Condition)
        }
        else{
            $Result = $DriverWait.Until[WebElement]($Condition)
        }        
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
    Write-Host "Creating edge driver."
    $EdgeService = [EdgeDriverService]::CreateDefaultService()
    $EdgeService.HideCommandPromptWindow = $true

    $EdgeOptions = New-Object EdgeOptions
    $EdgeOptions.AddArguments("headless")
    
    $EdgeDriver = New-Object EdgeDriver($EdgeService, $EdgeOptions) 
    
    $Detector = New-Object LocalFileDetector
    $EdgeDriver.FileDetector = $Detector

    $EdgeDriver.Navigate().GoToURL($StartUrl)

    Return $EdgeDriver
}

function SendKeysTo {
    [OutputType([WebElement])]
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$XPath,
        [string]$Keys,
        [bool]$IsHidden = $false
    )

    $ElementToSend = GetElementFromPath $EdgeDriver -XPath $XPath -IsHidden $IsHidden
    $ElementToSend.SendKeys($Keys)

    return $ElementToSend
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

    $TextboxPassword = SendKeysTo $EdgeDriver -XPath $XPath -Keys (ConvertFrom-SecureString -SecureString $Password -AsPlainText)
    $Password.Clear()
    $TextboxPassword.Submit()
}

function Login {
    param (
        [EdgeDriver]$EdgeDriver
    )
    Write-Host -NoNewline -ForegroundColor Blue "Enter the username: "
    $Username = Read-Host

    $null = SendKeysTo $EdgeDriver -XPath '//*[@id="username"]' -Keys $Username

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

function ClickElement {
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$XPath,
        [WebElement]$ElementToCheck
    )

    $ElementToClick = GetElementFromPath $EdgeDriver -XPath $XPath -ElementToCheck $ElementToCheck
    $ElementToClick.Click()
}

function DoubleClickElement {
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$XPath
    )

    $ActionCustom = New-Object Actions($EdgeDriver)
    $FolderElement = GetElementFromPath $EdgeDriver -XPath $XPath
    $ActionCustom.DoubleClick($FolderElement).Perform()
}

function GoToFolder {
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$FolderName
    )

    DoubleClickElement $EdgeDriver -XPath "//td[.//span[contains(text(), 'Folder - $FolderName')]]"

    Write-Host "Current folder: $FolderName"
}

function GoToRoot {
    param (
        [EdgeDriver]$EdgeDriver
    )

    ClickElement $EdgeDriver -XPath '//button[@title="My files"]'
    Write-Host "Current folder: Root"
}

function UploadFile {
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$FilePath
    )

    $null = SendKeysTo $EdgeDriver -XPath '//input[@type="file"][@multiple]' -Keys $FilePath -IsHidden $true
    $EdgeDriver.ExecuteScript("document.querySelectorAll('input[type=\'file\'][multiple]').forEach(element => element.value='');")

    Write-Host "Upload of $FilePath started."
}

function DeleteOldest {
    param (
        [EdgeDriver]$EdgeDriver,
        [string]$FileName,
        [int]$ArchivesToKeep
    )

    try {        
        $ZipRows = GetElementFromPath $EdgeDriver -XPath "//tr[.//span[contains(text(), 'File - application/zip - ${FileName}')]]" -ReturnArray $true
         
        if ($ZipRows.Count -gt $ArchivesToKeep) {
            $ToDelete = $ZipRows | Sort-Object -Property ComputedAccessibleLabel | Select-Object -First 1

            ClickElement $EdgeDriver -XPath ".//td[.//input[@type='checkbox']]" -ElementToCheck $ToDelete

            ClickElement $EdgeDriver -XPath "//button[@data-testid='toolbar-trash']"

            $FileDeletedName = (($ToDelete.Text  -split '\n')[0] -split ' - ')[2]
            Write-Host "File $FileDeletedName deleted."
        }

        Write-Host "$($ZipRows.Count) archives founds. Nothing deleted."
    } catch { Write-Host "Nothing to delete." }
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

function DeleteLocalFiles {
    param (
        [string[]]$FilePaths
    )
     ForEach ($FilePath in $FilePaths) {
        If (Test-Path $FilePath) {
            Remove-Item -Path $FilePath -Force
        }
    }

    Write-Host "Temp files cleared."
}

function CloseDriver {
    param (
        [EdgeDriver]$EdgeDriver
    )
    $EdgeDriver.Close()
    $EdgeDriver.Quit()

    Write-Host "Goodbye."
}

function WaitToUpload {
    param (
        [EdgeDriver]$EdgeDriver,
        [int]$MinutesToWait,
        [int]$SecondsToSleep
    )
    $IsComplete = $false

    $ItemStatuses = GetElementFromPath $EdgeDriver -XPath "//div[contains(@class, 'transfers-manager-list-item')][.//progress]" -ReturnArray $true

    $EllapsedSeconds = 0

    while(-Not $IsComplete)
    {
        if ($EllapsedSeconds -gt $MinutesToWait * 60) {
            break
        }

        $AllDone = $true

        for($i=0; $i -lt $ItemStatuses.Count; $i++)
        {
            $ItemStatus = $ItemStatuses[$i]
            $ItemCompleted = $ItemStatus.Text.Contains("Uploaded")
            $AllDone = $ItemCompleted -And $AllDone
            $UploadInfo = ($ItemStatus.Text -replace  "`n|`r" -split '.zip',2)[0] + ".zip"
            $ProgressValue = ($ItemStatus.Text -replace  "`n","" -replace "`r","=" -split '.zip',2)[1] -split '=' | Select-Object -first 3 | Join-String -Separator ' '
            $IdBar = $i + 10
            if ($ItemCompleted) {
                Write-Progress -Activity $UploadInfo -Id $IdBar -Status "$ProgressValue" -PercentComplete 100 -Completed 
            } else {
                Write-Progress -Activity $UploadInfo -Id $IdBar -Status "$ProgressValue" -PercentComplete -1 
            }            
        }

        $IsComplete =  $AllDone
        Start-Sleep -Seconds $SecondsToSleep
        $EllapsedSeconds = $EllapsedSeconds + $SecondsToSleep
    }

    ClickElement $EdgeDriver -XPath "//button[@data-testid='drive-transfers-manager:close']"
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

$null = Login $EdgeDriver

GoToFolder $EdgeDriver -FolderName $iCloudDrive
UploadFile $EdgeDriver -FilePath $iCloudZipPath
DeleteOldest $EdgeDriver -FileName $iCloudDrive -ArchivesToKeep $ArchivesToKeep

GoToRoot $EdgeDriver

GoToFolder $EdgeDriver -FolderName $OneDrive
UploadFile $EdgeDriver -FilePath $OneDriveZipPath
DeleteOldest $EdgeDriver -FileName $OneDrive -ArchivesToKeep $ArchivesToKeep

WaitToUpload $EdgeDriver -MinutesToWait $MinutesToWait -SecondsToSleep $SecondsToSleep

DeleteLocalFiles $iCloudZipPath, $OneDriveZipPath

CloseDriver $EdgeDriver