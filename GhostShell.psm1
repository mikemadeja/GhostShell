#Requires -Version 5.0
<#

.SYNOPSIS
A PowerShell wrapper for Send-MailMessage that utilizes a JSON file for HTML and other config files.

.DESCRIPTION
#NEED TO ADD#

.EXAMPLE
 Get-GhostShellVariables

.LINK
https://github.com/mikemadeja/GhostShell/blob/master/README.md

#>

$MODULE_PATH = "C:\Program Files\WindowsPowerShell\Modules"
$MODULE_FOLDER_NAME = "GhostShell"
$GLOBAL_JSON_FILE = "Config.json"
$GLOBAL_JSON_BACKUP_FILE = "Config.bkup.json"
$GLOBAL_JSON = $MODULE_PATH + "\" + $MODULE_FOLDER_NAME + "\" + $GLOBAL_JSON_FILE
$GLOBAL_JSON_BACKUP = $MODULE_PATH + "\" + $MODULE_FOLDER_NAME + "\" + $GLOBAL_JSON_BACKUP_FILE
$DEFAULT_SMTP_ENTRY = "smtp.domain.com"

#INTERNAL FUNCTIONS
Function Test-GhostShellModulePath {
    $registryPSModulePath = ([Environment]::GetEnvironmentVariable("PSModulePath", "Machine")) -split ";"
    #Write-Output -Verbose "$registryPSModulePath"
    #Write-Output -Verbose "$PSScriptRoot"
    Foreach ($regEntry in $registryPSModulePath) {
        $regEntry -match $MODULE_PATH
    }
}

Function Test-GhostShellJSONFilePath {
    Write-Verbose "Testing if $GLOBAL_JSON exists..."
    If ((Test-Path $GLOBAL_JSON) -eq $True) {
        Write-Verbose -Message "$GLOBAL_JSON_FILE exists!"
        Write-Verbose -Message "Checking if $GLOBAL_JSON_FILE has an updated SMTP Server entry of $DEFAULT_SMTP_ENTRY..."
        If (((Get-GhostShellVariables).GLOBAL.mail.smtpServer) -eq $DEFAULT_SMTP_ENTRY) {
            Write-Error -Message "$DEFAULT_SMTP_ENTRY exists, please update the smtpServer entry in $GLOBAL_JSON_FILE to match your SMTP environment" -ErrorAction "Stop"
        }
    }
    Else {
        Write-Error -Message "$GLOBAL_JSON doesn't exist, copy $GLOBAL_JSON_BACKUP_FILE as $GLOBAL_JSON_FILE and update the parameters to fit your environment" -ErrorAction "Stop"
    }
}

Function Test-GhostShellJSONFile {
  Try {
    (Get-Content $GLOBAL_JSON | ConvertFrom-Json) | Out-Null
  }
  Catch {
    Write-Error "Invalid $GLOBAL_JSON_FILE file!" -ErrorAction "Stop"
  }
}

Function New-RandomString {
    (-join ((48..57) + (97..122) | Get-Random -Count 32 | ForEach-Object {[char]$_})).ToUpper()
}
Function Get-GhostShellMailMessageOptionalParameters {   
    [hashtable]$hashtable = @{"Bcc" = $Bcc;}
    Write-Output $OptionalParam
    }

Function Create-HTMLFormat {
     #Prepare HTML code
     
     $CSS = (Get-GhostShellVariables).GLOBAL.html.css
     $PostContent = "<br><a href='https://github.com/mikemadeja'>GhostShell on GitHub</a>"
     $HTMLParams = @{
        'Head' = $CSS;
        'PostContent' = $PostContent;
    }
    Write-Output $HTMLParams
}

Function Create-HTMLFragments {
    $Fragments = @()
    $Logo = (Get-GhostShellVariables).GLOBAL.html.logo
    $H1 = "<H1><Img src=$Logo width=`"64`" height=`"64`">&nbsp;&nbsp;$Header</H1><br>"
    If ($HttpComment -ne $null) {
        $H3 = "<H3><a href='$HttpLink'>$HttpComment</a></H3>"
    }
    $Fragments = $H1 + $H3
    Write-Output $Fragments
}

#EXTERNAL FUNCTIONS
Function Get-GhostShellVariables {
<#

.SYNOPSIS
A PowerShell wrapper for Send-MailMessage that utilizes a JSON file for HTML and other config files.

.DESCRIPTION
#NEED TO ADD#

.EXAMPLE
Get-GhostShellVariables
(Get-GhostShellVariables).GLOBAL.mail
(Get-GhostShellVariables).GLOBAL.mail.smtpServer
(Get-GhostShellVariables).GLOBAL.html
(Get-GhostShellVariables).GLOBAL.html.css

.LINK
https://github.com/mikemadeja/GhostShell/blob/master/README.md

#>
    Test-GhostShellJSONFile
    (Get-Content $GLOBAL_JSON | ConvertFrom-Json)
}
Function Send-GhostShellMailMessage {
    Param(
        [Parameter(Mandatory=$true)]
        $To,
        [Parameter(Mandatory=$true)]
        [String]$Subject,
        [Parameter(Mandatory=$true)]
        [String]$Header,
        [parameter(Mandatory=$false,ParameterSetName = 'Http')]
        [ValidateNotNullorEmpty()]
        [String]$HttpComment,
        [Parameter(Mandatory=$false,ParameterSetName = 'Http')]
        [ValidateNotNullorEmpty()]
        [String]$HttpLink,
        [Parameter(Mandatory=$true)]
        $Body,
        [Parameter(Mandatory=$true)]
        [String]$From,
        $Bcc,
        $Cc,
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty,
        [Parameter()]
        [ValidateNotNullorEmpty()]
        [Switch]$AttachAsHTML,
        [Parameter()]
        [ValidateNotNullorEmpty()]
        [String]$AttachAsHTMLFileName,
        [Switch]$AttachAsPDF,
        [Int]$Port = 25,
        [Hashtable]$OptionalParameters
    )
    #Test to make sure the JSON file is valid
    Test-GhostShellJSONFilePath

    #Pull SMTP server from JSON file
    $SmtpServer = (Get-GhostShellVariables).GLOBAL.mail.smtpServer

    #Call function to create HTML format
    $HTMLParams = Create-HTMLFormat
    #Call function to construct the HTML data
    $Fragments = Create-HTMLFragments
    
    Write-Verbose -Message $HTMLParams
    Write-Verbose -Message $Fragments
    Write-Verbose -Message "Checking Body if it's a string or not"
    If (($Body.GetType().Name -eq "String")) {
        Write-Verbose -Message "Body is string"
        $Fragments += $Body
        $Body = ConvertTo-Html -PreContent ($Fragments | Out-String) @HTMLParams | Out-String
    }
    Else {
        Write-Verbose -Message "Body is not string"
        $Fragments += $Body | ConvertTo-Html
        $Body = ConvertTo-Html -Body ($Fragments | Out-String) @HTMLParams | Out-String
    }

    If (($AttachAsHTML.IsPresent) -eq $true){
        If ($AttachAsHTMLFileName -eq $null) {
            $HTMLOutputFile = 'C:\Users\Mike\AppData\Local\Temp\\' + ((New-RandomString) + ".HTML").ToString()
            $Body | Out-File $HTMLOutputFile
        }
    }

    $DefaultSmtpParams = @{
        'SmtpServer' = $SmtpServer;
        'To' = $To;
        'From' = $From;
        'Subject' = $Subject;
        'Body' = $Body;
        'Port' = $Port;
    }
    Write-Verbose -Message "Default SMTP Parameters"
    Write-Verbose -Message $DefaultSmtpParams
    $OptionalParameters = Get-GhostShellMailMessageOptionalParameters

    If ($OptionalParameters -ne $null) {
        Send-MailMessage @DefaultSmtpParams @OptionalParameters -BodyAsHtml -Credential $Credential -UseSsl
    }
    Else {
        Send-MailMessage @DefaultSmtpParams -BodyAsHtml -Credential $Credential -UseSsl
    }
    
}

Export-ModuleMember -Function Get-GhostShellVariables
Export-ModuleMember -Function Send-GhostShellMailMessage
