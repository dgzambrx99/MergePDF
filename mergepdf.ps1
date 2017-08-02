# Programmer: Luka Ivicevic
# Client: David Zambrano
# Description: This script will navigate to the given directory and combine pdfs in each folder in that directory one level deep, then delete the processed pdfs.
# Run with: mergepdf -path C:\mypath
# v2 creates and sends an email indicating while files has been processed or indicating there are no files to process - DZ
# Added logging to a file and says who the email is going to - DZ
# Replaced Write-Host with Write-Output - DZ
# Added excluded files to the email

# Get command line args - removed variable, can't get it to work in task scheduler -DZ
# Param([string]$path)

Param (
	[string]$Path = "\\afs1\ScannedInvoices",
	[string]$PrintPath = "pathToFolderYouWantToPrint",
	[string[]]$Exclude = @('Space Exploration Technologies (Spacex)'),
	[string]$SMTPServer = "192.187.218.190",
	[string]$From = "mailbox@astro.local",
	[string[]]$To = @('Mark Streiff<mstreiff@astro.local>', 'Josh Hansen<jhansen@astro.local>'),
	[string]$Subject = "Invoice File List"
	)

# Starts the logfile - DZ
Start-Transcript -path 'C:\Users\administrator.ASTRO\Documents\InvoiceProcessor\logfile.log' -append

# Import libraries.
Add-Type -Path $PSScriptRoot\Libraries\PDFsharp\code\PdfSharp\bin\Debug\PdfSharp.dll

# This function merges all pdfs in myfolderpath and saves it to foldername, then deletes the processed pdfs.
# Merge-PDF "FullSystemPath" "FileName"
Function Merge-PDF($myfolderpath, $foldername) {
	# Create output pdf and open PdfSharp objects.
	$output = New-Object PdfSharp.Pdf.PdfDocument
	$PdfReader = [PdfSharp.Pdf.IO.PdfReader]
	$PdfDocumentOpenMode = [PdfSharp.Pdf.IO.PdfDocumentOpenMode]

	# Merge PDFs into one
	foreach($i in (gci $myfolderpath *.pdf | Where-Object { $_.CreationTime.toString("MM/dd/yyyy") -ceq (Get-Date -Format MM/dd/yyyy) })) {
		try {
			Write-Output $i.Name
			$input = $PdfReader::Open($i.FullName, $PdfDocumentOpenMode::Import)
			$input.Pages | %{$output.AddPage($_)}
			# Delete processed pdfs.
			Remove-Item($i.FullName)
		} catch {
			Write-Output "Error occurred while processing file: " $i.FullName ", skipping this file."
			Write-Output "ERROR: " + $_.Exception.Message + " Failed Item: " + $_.Exception.ItemName
		}
	}
	try {
		# Create file output path
		$fileoutput = $myfolderpath + "\" + $foldername + ".pdf"
		$output.Save($fileoutput)
	} catch {
		Write-Output "Error occurred while saving file: " $fileoutput ", skipping this file."
		Write-Output "ERROR: " $_.Exception.Message " Failed Item: " $_.Exception.ItemName
	}
	
}

#Compose an e-mail - DZ
$SMTPMessage = @{
    To = $To
    From = $From
	Subject = "$Subject at $Path"
    Smtpserver = $SMTPServer
}

# Iterate through each folder in $path alphabetically.v2 sort name? -DZ
Get-ChildItem -Exclude $Exclude -Path $Path | Sort-Object name | ?{$_.PSIsContainer -and $_.GetFiles("*.pdf").Count} | Select-Object FullName | % {
	Write-Host "`nFullPath: " $_.FullName
	$filepath = $_.FullName
	$filename = Split-Path $_.FullName -Leaf
	$filename = $filename + " $(Get-Date -Format MM-dd-yyyy)"
	Merge-PDF $filepath $filename
} | Sort-Object $_.FullName

# v2 Combines pdfs in print path and prints the combined pdf. -LI
$filename = Split-Path $PrintPath -Leaf
$filename = $filename + " $(Get-Date -Format MM-dd-yyyy)"
Merge-PDF $PrintPath $filename
$filepath = $PrintPath + "\" + $filename + ".pdf"
Write-Host "Printing file $($filepath)..."
Start-Process -FilePath $filepath -Verb Print

# v2 Creates a clickable e-mail of folders with pdfs in the directory, name sort?? -DZ
$File = Get-ChildItem -Exclude $Exclude -Path $Path | Sort-Object name | ?{$_.PSIsContainer -and $_.GetFiles("*.pdf").Count}
If ($File)
{	$SMTPBody = "`r`nFiles Excluded:`r`n$Exclude <br>`n<p>`nThe following files have recently been added/changed:<br>`n<p>`n"
	$File | ForEach { $SMTPBody += "<td>""$($_.FullName)""</td><br>`n" }
	Send-MailMessage @SMTPMessage -BodyAsHtml $SMTPBody -Encoding Unicode
    Write-Output("`r`nSending report E-mail to:`r`n$To")
}
If (-Not $File) # Sends an e-mail if there are no files to process indicating there are no files to process
{   $SMTPBody = "`r`nNo files to process<br>`n<p>`n"
    Send-MailMessage @SMTPMessage -BodyAsHtml $SMTPBody -Encoding Unicode
    Write-Output("`r`nSending no op E-mail to:`r`n$To")
}
# Displays while files are excluded from processing - DZ
Write-Output("`r`nFiles Excluded:`r`n$Exclude")

Write-Output("`r`nDone.`r`n")

Stop-Transcript