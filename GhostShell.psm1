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
#EXTERNAL FUNCTIONS
Function Get-GhostShellVariables {
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
    Test-GhostShellJSONFilePath
    #Prepare HTML code
    $Logo = (Get-GhostShellVariables).GLOBAL.html.logo
    $CSS = (Get-GhostShellVariables).GLOBAL.html.css
    $SmtpServer = (Get-GhostShellVariables).GLOBAL.mail.smtpServer
    $H1 = "<H1><Img src=$Logo width=`"64`" height=`"64`">&nbsp;&nbsp;$Header</H1><br>"
    If ($HttpComment -ne $null) {
        $H3 = "<H3><a href='$HttpLink'>$HttpComment</a></H3>"
    }

    $PostContent = "<br><a href='https://github.com/mikemadeja'>GhostShell on GitHub</a>"
    $Fragments = @()
    $Fragments = $H1 + $H3

    $HTMLParams = @{
        'Head' = $CSS;
        'PostContent' = $PostContent;
    }
    
    If (($Body.GetType().Name -eq "String")) {
        $Fragments += $Body
        $Body = ConvertTo-Html -PreContent ($Fragments | Out-String) @HTMLParams | Out-String
    }
    Else {
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
