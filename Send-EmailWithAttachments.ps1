<#  
.SYNOPSIS   
     Send email containing multiple attachments using secure SMTP 
.DESCRIPTION  

This runbook send email with multiple attachments through secure SMTP. The attached files are 
constructed in-memory which eliminates the need of using workarounds like leveraging ‘temp’ 
folder on sandbox machine executing runbook to copy the files first and then use full path 
to files with ‘Send-EmailMessage’ cmdlet. The common use case of this runbook is to attach 
files predominately containing text data (example: csv,txt etc files). 

Prereq:     
    1. Automation PS Credential Asset with userid and password to smtp service. 
    
.PARAMETER Source  
    String email address that will be used as sender
    Example: no-reply@domain.com 
  
.PARAMETER To 
    String email address of destination  
    Example: target@domain.com

.PARAMETER Server 
    String SMTP server address
    Example: smtp.live.com

.PARAMETER Port 
    String SMTP server port number
    Example: 587
         
.PARAMETER Subject
    String Subject of mail message
    
.PARAMETER fileNames
     Multi-Dimensional Array (String) - Contain name of attached files
     Example: ["Data.csv","ReadMe.txt"]

.PARAMETER fileContents
     Multi-Dimensional Array (String) - Content attach files
     Example: ["vmsize,vmname,servicename","The attached file contain list of vms"]

.PARAMETER credName 
    String - Name of PS Credential asset
    Example: SMTPCredential


.EXAMPLE  
  Send-EmailWithAttachments "no-reply@domain.com" "target@domain.com" "smtp.live.com" "587" "[VM-GOVERNENCE] List of non-iaas vms" "PFA" $CsvFileNames $csvFileContents (Get-Credential)

.NOTES  
    Author: Razi Rais  
    Website: www.razibinrais.com
    Last Updated: 4/10/2015   
#>

workflow Send-EmailWithAttachments
{
       param ([Parameter(Mandatory=$true,Position=0)][string]$source,
       [Parameter(Mandatory=$true,Position=1)][string]$to,
       [Parameter(Mandatory=$true,Position=2)][string]$server,
       [Parameter(Mandatory=$true,Position=3)][string]$port,
       [Parameter(Mandatory=$true,Position=4)][string]$subject, 
       [Parameter(Mandatory=$true,Position=5)][string]$body,
       [Parameter(Mandatory=$true,Position=6)][string[]]$fileNames,
       [Parameter(Mandatory=$true,Position=7)][string[]]$fileContents, 
       [Parameter(Mandatory=$true,Position=8)][string]$credName 
              
       )


$cred = Get-AutomationPSCredential -Name $credName 

  inlinescript 
  {
    $strFiles =  $Using:fileContents;  
    $message = New-Object System.Net.Mail.MailMessage
    $smtpClient = New-Object System.Net.Mail.smtpClient($Using:server,$Using:port)
    $smtpClient.EnableSsl = $true
    $smtpClient.Credentials = $Using:cred 
    $recipient = New-Object System.Net.Mail.MailAddress($Using:to, "Recipient") 
    $sender = New-Object System.Net.Mail.MailAddress($Using:source) 
    $message.Sender = $sender
    $message.From = $sender
    $message.Subject = $Using:subject
    $message.To.add($recipient)
    $message.Body = $Using:body
    
        $count = 0;
        foreach($fileName in $Using:fileNames)
        {
            # Prepare file attachement. Create a memory stream
            $memoryStream = New-Object IO.MemoryStream
            [Byte[]]$contentAsBytes = [Text.Encoding]::UTF8.GetBytes($strFiles[$count])
            $memoryStream.Write($contentAsBytes, 0, $contentAsBytes.Length)
            # Set the position to the beginning of the stream.
            [Void]$memoryStream.Seek(0, 'Begin')
            # Create attachment
            $contentType = New-Object Net.Mime.ContentType -Property @{
              MediaType = [Net.Mime.MediaTypeNames+Application]::Octet
              Name = $fileName
            }
            $attachment = New-Object Net.Mail.Attachment $memoryStream, $contentType
    
            # Add the attachment
            $message.Attachments.Add($attachment)
            
            $count++;
         }
    # Send Mail via SmtpClient
    $smtpClient.Send($message)
  }
}