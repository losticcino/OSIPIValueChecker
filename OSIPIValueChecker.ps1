#
# The intent of this script is to allow scheduled alerting of a pi value from a client connection
#
param([switch]$UpdateCredentials,[switch]$UpdateServerConfig,[switch]$UpdateRecipients,[switch]$UpdateSiteName)

# program variables - don't modify these.
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.MessageBox") | Out-Null
try {$emailCredentials = Import-Clixml .\Email.credentials}
catch {$UpdateCredentials = $true}
try {$emailServer = Import-Clixml .\Email.server}
catch {$UpdateServerConfig = $true}
try {$tags = Import-Csv .\taglist.csv}
catch {$UpdateTagList = $true}
try {$recipients = Get-Content .\Recipient.list -ErrorAction Stop}
catch {$UpdateRecipients = $true}
try {$siteName = Get-Content .\Site.name -ErrorAction Stop}
catch {$UpdateSiteName = $true}
try {$lastMessage = Import-Clixml .\LastMessage.date}
catch {$InitializeLastMessage = $true}

# functions

Function FindValueIndex ($result) {
	for ($i = 0; $i -le $result.Count; $i++) {
		if ($result[$i].Contains("Snapshot value")) {
			for ($j = $i; $j -le $result.Count; $j++) {
				if ($result[$j].Contains("Value =")) {
					return $j
				}
			}
		}
	}
	return -1
}

Function SendEmail ($Server, $Credentials, $Destination, $Subject, $Message) {
	if (($Server -eq $null) -or ($Credentials -eq $null)) {exit}
	if ($Destination -eq $null) {$Destination = $Credentials.UserName}
	if ($Subject -eq $null) {$Subject = "Scripted Notification from " + $Credentials.UserName}
	if ($Message -eq $null) {$Message = "An alert was generated, but a message was not specified."}
	$Destination = $Destination.split(',')
	$Destination = $Destination.split(';')
	$messageParams = @{
		SmtpServer = $Server.SmtpServer
		Port = $Server.Port
		UseSSL = $Server.UseSSL
		Credential = $Credentials
		From = $Credentials.UserName
		To = $Destination
		Subject = $Subject
		Body = $Message
	}

	try {
		Send-MailMessage @messageParams -ErrorAction Stop
		$messageParams.Add("Result","Success")
	}
	catch {$messageParams.Add("Result",$Error[0])}
	return $messageParams
}

Function UpdateServerConfig ($Config) {
	if ($Config -eq $null) {
		$Config = @{
			SmtpServer = ''
			Port = 587
			UseSSL = $true
			Credential = ''
			From = ''
			To = ''
			Subject = ''
			Body = ''
		}
	}
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Visualbasic") | Out-Null
	
	# Update Server
	$svr = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Server Address","Email Server Config",$emailServer.SmtpServer)
	if (!($svr -eq '')) {$Config.SmtpServer = $svr}

	#update Port
	[int]$prt = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Server Port","Email Server Config",$emailServer.Port)
	if (($prt -gt 1) -and ($prt -lt 65535)) {$Config.Port = $prt}

	# Update SSL option
	$ssl = [System.Windows.Forms.MessageBox]::Show("Use TLS/SSL?","Email Server Config","YesNo")
	$Config.UseSSL = switch ($ssl) {Yes{$true} No{$false}}

	return $Config
}

Function UpdateServerCredentials ($Credentials) {
	if ($Credentials -eq $null) {$Credentials = ""}
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Visualbasic") | Out-Null
	[System.Windows.Forms.MessageBox]::Show("You must enter and verify your email credentials to continue","OK") | Out-Null
	
	#verify passwords match
	$i = 0
	do {
		$pass = $false
		$cred1 = Get-Credential -Message "Enter your email server credentials"
		$cred2 = Get-Credential -Message "Verify your credentials" -UserName $cred1.UserName
		if ([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR((($cred1.Password)))) -eq [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR((($cred2.Password))))) {
			$i = 100
			$Credentials = $cred1
		}
		else {
			[System.Windows.Forms.MessageBox]::Show("Passwords did not match, try again.","Credential Error","OK") | Out-Null
			$i++
		}
	} while ($i -lt 5)

	if ($i -le 10) {
		[System.Windows.Forms.MessageBox]::Show("Verification failed "+$i+" times and cannot contiue.  Please verify your keyboard operation and try again.","Credentials Entry Error","OK") | Out-Null
		exit
	}
	elseif ($i -eq 100) {
		$messagePopup = New-Object -ComObject wscript.shell -ErrorAction Stop
		$messagePopup.Popup("Testing server connectivity",5,"Email Setup",0)
		if (!(Test-Connection $emailServer.SmtpServer -Count 1)) {
			[System.Windows.Forms.MessageBox]::Show("Cannot connect to server " + $emailServer.SmtpServer +"`nVerify network connectivity and server configuration`nExiting...","Email Server Error","OK")
			return $Credentials
		}
		$destination = [Microsoft.VisualBasic.Interaction]::InputBox("What email address should I send a test message to?","Test Email Destination",$emailCredentials.Username)
		$test = SendEmail -Server $emailServer -Credentials $Credentials -Destination $destination -Subject "Email Test Message" -Message "This should indicate a successful test message."
		
		if (!($test.Result -like "Success")) {[System.Windows.Forms.MessageBox]::Show("Test Email Failed.  Results are: `n"+$test.Result,"Email Server Error","OK")}
	}
	return $Credentials
}

Function GenerateSamepleTaglist {
	"Server,TagName,FriendlyName,Threshold`nserver,tag,Tag,1`nserver,tag,Tag,2" | Out-File taglist.csv -Encoding ascii
}

Function UpdateRecipients ($Recipients) {
	if ($Recipients -eq $null) {$Recipients = "sample1@email.com,sample2@email.com"}
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Visualbasic") | Out-Null
	$Recipients = [Microsoft.VisualBasic.Interaction]::InputBox("Enter recipient email addresses, with commas separating additional recipients.","Email Server Config",$Recipients)
	$Recipients.Replace(';',',')
	return $Recipients
}

Function UpdateSiteName ($SiteName) {
	if ($SiteName -eq $null) {$SiteName = ""}
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Visualbasic") | Out-Null
	$SiteName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Site Name","Update Site Name",$SiteName)
	return $SiteName
}

# main program

if ($UpdateServerConfig) {
	$emailServer = UpdateServerConfig $emailServer
	$emailServer | Export-Clixml .\Email.server
}

if ($UpdateCredentials) {
	$emailCredentials = UpdateServerCredentials $emailCredentials
	$emailCredentials | Export-Clixml .\Email.credentials
}

if ($UpdateTagList) {
	GenerateSamepleTaglist
	[System.Windows.Forms.MessageBox]::Show("A Sample tag list has been generated, and will be opened for editing at this time.","OK") | Out-Null
	Start-Process notepad.exe -ArgumentList .\taglist.csv
	[System.Windows.Forms.MessageBox]::Show("Please click OK when you have finished editing the tag list.","OK") | Out-Null
	try {$tags = Import-Csv .\taglist.csv}
	catch {[System.Windows.Forms.MessageBox]::Show($Error[0],"An Error Occurred","OK") | Out-Null}
}

if ($UpdateSiteName) {
	$siteName = UpdateSiteName $siteName
	$siteName | Out-File .\Site.name -Encoding ascii -ErrorAction SilentlyContinue
}

if ($UpdateRecipients) {
	$recipients = UpdateRecipients $recipients
	if ([System.Windows.Forms.MessageBox]::Show("Send Test Email?","Updated Recipients","YesNo") -eq 'Yes') {
		SendEmail -Server $emailServer -Credentials $emailCredentials -Destination $recipients -Subject ("Test Message from " + $siteName) -Message ("This is a test message for alerts from " + $siteName)
	}
	$recipients | Out-File .\Recipient.list -Encoding ascii -ErrorAction SilentlyContinue
}

if ($InitializeLastMessage) {
	$lastMessage = @{}
	foreach ($tag in $tags) {
		$lastMessage.Add($tag.TagName,[datetime]0)
	}
}

foreach ($tag in $tags) {
	$result = apisnap $tag.Server $tag.TagName

	#Find Snapshot Value (Just in case it ever moves index)
	$valueIndex = FindValueIndex $result
	$startPos = $result[$valueIndex].LastIndexOf("Value = ") + 8
	$valLen = ($result[$valueIndex].IndexOf(" ",($result[$valueIndex].LastIndexOf("Value = ") + 8)) - ($result[$valueIndex].LastIndexOf("Value = ") + 8))
	[float]$piValue = $result[$valueIndex].Substring($startPos,$valLen)
	if (($piValue -ge $tag.threshold) -and ($lastMessage.($tag.TagName) -eq [datetime]0)) {
		$strSubject = ("Value Threshold Exceeded for " + $tag.FriendlyName + " at " + $siteName + ".")
		$strMessage = ("The tag " + $tag.TagName + " exceeded the threshold of " + $tag.Threshold + ".`n`nThe PI snapshot value is as follows: " + $result[$valueIndex])
		$sent = SendEmail -Server $emailServer -Credentials $emailCredentials -Destination $recipients -Subject $strSubject -Message $strMessage
		if ($sent.Result -like "Success") {
			$lastMessage.($tag.TagName) = [datetime](Get-Date)
			$lastMessage | Export-Clixml .\LastMessage.date
		}
	}
	elseif (($piValue -lt ([float]$tag.threshold * 0.9)) -and ($lastMessage.($tag.Tagname) -lt (Get-Date).AddMinutes(-5))) {
		if ($lastMessage.($tag.TagName) -gt [datetime]0) {
			$lastMessage.($tag.TagName) = [datetime]0
			$lastMessage | Export-Clixml .\LastMessage.date
		}
	}
}