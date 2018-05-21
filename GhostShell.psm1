#Requires -Version 5.0
<#

.DESCRIPTION
The GhostShell module enables you to have a single configuration file to send unified emails.

.LINK
https://github.com/mikemadeja/GhostShell/wiki

#>
Function New-RandomString {
    (-join ((48..57) + (97..122) | Get-Random -Count 32 | ForEach-Object {[char]$_})).ToUpper()
}

$MODULE_PATH = "C:\Program Files\WindowsPowerShell\Modules"
$MODULE_FOLDER_NAME = "GhostShell"
$GLOBAL_JSON_FILE = "Config.json"
$GLOBAL_JSON_BACKUP_FILE = "Config.bkup.json"
$GLOBAL_JSON = $MODULE_PATH + "\" + $MODULE_FOLDER_NAME + "\" + $GLOBAL_JSON_FILE
$GLOBAL_JSON_BACKUP = $MODULE_PATH + "\" + $MODULE_FOLDER_NAME + "\" + $GLOBAL_JSON_BACKUP_FILE
$DEFAULT_SMTP_ENTRY = "smtp.domain.com"
$ENV_TEMP = $ENV:Temp
$ENV_PATHS = $ENV:PATH -split ";"
$PDF_APPLICATION = "wkhtmltopdf.exe"
$RANDOM_STRING = New-RandomString
$RANDOM_FILE_NAME_HTML = ($RANDOM_STRING + ".HTML").ToString()
$RANDOM_FILE_NAME_PDF = ($RANDOM_STRING + ".PDF").ToString()
$TEMP_FILE_HTML = $ENV_TEMP + "\" + $RANDOM_FILE_NAME_HTML
$TEMP_FILE_PDF = $ENV_TEMP + "\" + $RANDOM_FILE_NAME_PDF
#INTERNAL FUNCTIONS
Function Test-GhostShellModulePath {
    $registryPSModulePath = ([Environment]::GetEnvironmentVariable("PSModulePath", "Machine")) -split ";"
    Foreach ($regEntry in $registryPSModulePath) {
        $regEntry -match $MODULE_PATH
    }
}

Function Test-PDFApplication {
    $ENVCount = 1
    Foreach ($ENVPath in $ENV_PATHS) {
        $ENVPathsCount = $ENV_PATHS.Length
        
        $PdfApplicationFullPath = $ENVPath + "\" + $PDF_APPLICATION
        If ((Test-Path $PdfApplicationFullPath) -eq $False) {
            If ($ENVPathsCount -eq $ENVCount) {
                Write-Error -Message "Cannot find $PDF_APPLICATION, please make sure the $PDF_APPLICATION is installed and part of the Environment Variables for PATH" -ErrorAction "Stop"
            }
            $ENVCount++
        }
        Else {
            Write-Verbose -Message "Found $PDF_APPLICATION"
        }
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

Function Test-HTTPComment {
    If ($HttpComment -ne $null) {
        Return $True
    }
    Else {
        Return $False
    }
}

Function Test-HTTPLink {
    If ($HTTPLink -match "http*://") {
        Return $True
    }
    Else {
        Return $False
    }
}

Function Get-GhostShellMailMessageOptionalParameters {
    $OptionalParams = @{}
    if (($UseSSL.IsPresent) -eq $True) {
        Write-Verbose -Message "UseSSL is present"
        $OptionalParams  += @{"UseSSL" = $True;}
    }
    if ($Credential -ne $null) {
         Write-Verbose -Message "Credential is present"
        $OptionalParams  += @{"Credential" = $Credential;}
    }
    if ($Bcc -ne $null) {
        Write-Verbose -Message "Bcc is present"
        $OptionalParams  += @{"Bcc" = $Bcc;}
    }
    if ($Cc -ne $null) {
        Write-Verbose -Message "Cc is present"
        $OptionalParams  += @{"Cc" = $Cc;}
    }

    if ($Attachments) {
        $allAttachments += $Attachments
    }
    if ($AttachBodyAsPDF.IsPresent) {
        $allAttachments += $TEMP_FILE_PDF
    }
    if ($AttachBodyAsHTML.IsPresent) {
        $allAttachments += $TEMP_FILE_HTML
    }

    if ($allAttachments -ne $null) {
        $OptionalParams += @{"Attachments" = $allAttachments;}
    }
    Write-Output $OptionalParams
}
Function ConvertTo-GhostShellHTML ($RANDOM_FILE_NAME_HTML) {
    $Body | Out-File $TEMP_FILE_HTML
}
Function ConvertTo-GhostShellPDF {
    Test-PDFApplication
    If ((Test-Path -Path $TEMP_FILE_HTML) -eq $False) {
        $Body | Out-File $TEMP_FILE_HTML
    }
    $Application = "wkhtmltopdf.exe"
    $Quiet = "-q"
    #$DPI = "-d 600 "
    #&$Application $Quiet $DPI $TEMP_FILE_HTML $TEMP_FILE_PDF
    &$Application @($Quiet, "--disable-smart-shrinking", "-d", "600", $TEMP_FILE_HTML, $TEMP_FILE_PDF)
}
Function Create-HTMLFormat {
     #Prepare HTML code
     
     $CSS = (Get-GhostShellVariables).GLOBAL.html.css
     $PostContent = "<br><a href='https://github.com/mikemadeja/GhostShell' target=`"_blank`">GhostShell on GitHub</a>"
     $HTMLParams = @{
        'Head' = $CSS;
        'PostContent' = $PostContent;
        'Title' = $Subject;
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

Function Remove-TempFiles  {
    If ((Test-Path -Path $TEMP_FILE_HTML) -eq $True) {
        Remove-Item $TEMP_FILE_HTML -Force
    }
    If ((Test-Path -Path $TEMP_FILE_PDF) -eq $True) {
        Remove-Item $TEMP_FILE_PDF -Force
    }
}

#EXTERNAL FUNCTIONS
Function Get-GhostShellVariables {
<#

.DESCRIPTION
Get-GhostShellVariables enables you to call Config.json for global variables that can be used in your scripts. Please use Get-Help Get-GhostShellVariables -Online for more info.

.LINK
https://github.com/mikemadeja/GhostShell/wiki/GhostShell-Commands#get-ghostshellvariables

#>
    Test-GhostShellJSONFile
    (Get-Content $GLOBAL_JSON | ConvertFrom-Json)
}

Function Send-GhostShellMailMessage {
    <#
    
    .DESCRIPTION
    Send-GhostShellMailMessage enables you to send a unified email that utilizes the Config.json file. Please use Get-Help Send-GhostShellMailMessage -Online for more info.
    
    .LINK
    https://github.com/mikemadeja/GhostShell/wiki/GhostShell-Commands#send-ghostshellmailmessage
    
    #>
    Param(
        [Parameter(Mandatory=$true)]
        $To,
        [Parameter(Mandatory=$true)]
        [String]$Subject,
        [String]$Header,
        [parameter(Mandatory=$false,ParameterSetName = 'Http')]
        [ValidateNotNullorEmpty()]
        [String]$HttpComment,
        [Parameter(Mandatory=$false,ParameterSetName = 'Http')]
        [ValidateNotNullorEmpty()]
        [String]$HttpLink,
        [Parameter(Mandatory=$true)]
        $Body,
        [String]$From,
        $Bcc,
        $Cc,
        [String[]]$Attachments,
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential,
        [Parameter()]
        [ValidateNotNullorEmpty()]
        [Switch]$AttachBodyAsHTML,
        [Parameter()]
        [ValidateNotNullorEmpty()]
        [Switch]$AttachAsHTML,
        [Parameter()]
        [ValidateNotNullorEmpty()]
        [String]$AttachBodyAsHTMLFileName,
        [Switch]$AttachBodyAsPDF,
        [Parameter()]
        [ValidateNotNullorEmpty()]
        [Switch]$AttachAsPDF,
        [Parameter()]
        [ValidateNotNullorEmpty()]
        [Switch]$UseSSL,
        [Int]$Port = 25
    )
    #Test to make sure the JSON file is valid
    Test-GhostShellJSONFilePath
<#     If ((Test-HTTPLink) -eq $True -and (Test-HTTPComment) -eq $False) {
        Write-Error "Missing HTTPComment parameter" -ErrorAction "Stop"
    }
    If ((Test-HTTPLink) -eq $False -and (Test-HTTPComment) -eq $True) {
        Write-Error "Missing HTTPLink parameter" -ErrorAction "Stop"
    } #>

    #Call function to create HTML format
    $HTMLParams = Create-HTMLFormat
    #Call function to construct the HTML data
    $Fragments = Create-HTMLFragments
    
    Write-Verbose -Message $HTMLParams
    Write-Verbose -Message $Fragments
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
    If (($AttachBodyAsHTML.IsPresent) -eq $True){
        Write-Verbose -Message "AttachBodyAsHTML is $($AttachBodyAsHTML.IsPresent)"
        ConvertTo-GhostShellHTML $TEMP_FILE_HTML
        Write-Verbose -Message $TEMP_FILE_HTML
    }
    If (($AttachBodyAsPDF.IsPresent) -eq $True){
        Write-Verbose -Message "AttachBodyAsPDF is $($AttachBodyAsPDF.IsPresent)"
        ConvertTo-GhostShellPDF $TEMP_FILE_PDF
        Write-Verbose -Message $TEMP_FILE_PDF
    }

    $DefaultSmtpParams = @{
        'SmtpServer' = (Get-GhostShellVariables).GLOBAL.mail.smtpServer;
        'To' = $To;
        'From' = (Get-GhostShellVariables).GLOBAL.mail.smtpFrom;
        'Subject' = $Subject;
        'Body' = $Body;
        'Port' = $Port;
    }
    Write-Verbose -Message "Default SMTP Parameters"
    Write-Verbose -Message $DefaultSmtpParams
    $OptionalParameters = Get-GhostShellMailMessageOptionalParameters
    If ($OptionalParameters -ne $null) {
        Write-Verbose -Message "Sending SMTP with optional parameters"
        Send-MailMessage @DefaultSmtpParams @OptionalParameters -BodyAsHtml
    }
    Else {
        Write-Verbose -Message "Sending SMTP without optional parameters"
        Send-MailMessage @DefaultSmtpParams -BodyAsHtml
    }

    Remove-TempFiles
    If ($AttachAsHTML.IsPresent) {
        Write-Warning -Message "AttachAsHTML is depricated, please use AttachBodyAsHTML!"
    }
        If ($AttachAsPDF.IsPresent) {
        Write-Warning -Message "AttachAsPDF is depricated, please use AttachBodyAsHTML!"
    }
}

Export-ModuleMember -Function Get-GhostShellVariables
Export-ModuleMember -Function Send-GhostShellMailMessage
