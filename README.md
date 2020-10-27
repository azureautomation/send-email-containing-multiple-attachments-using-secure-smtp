Send email containing multiple attachments using secure SMTP
============================================================

            

**Description**

This runbook send email with multiple attachments through secure SMTP. The attached files are

constructed in-memory which eliminates the need of using workarounds like leveraging ‘temp’

folder on sandbox machine executing runbook to copy the files first and then use full path

to files with ‘Send-EmailMessage’ cmdlet. The common use case of this runbook is to attach

files predominately containing text data (example: csv,txt etc files).

 

 

        
    
TechNet gallery is retiring! This script was migrated from TechNet script center to GitHub by Microsoft Azure Automation product group. All the Script Center fields like Rating, RatingCount and DownloadCount have been carried over to Github as-is for the migrated scripts only. Note : The Script Center fields will not be applicable for the new repositories created in Github & hence those fields will not show up for new Github repositories.
