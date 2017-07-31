# Programmer: Luka Ivicevic
# Client: David Zambrano
# Description: This script will navigate to the given directory and combine pdfs in each folder in that directory one level deep, then delete the processed pdfs.
# Run with: mergepdf -path C:\mypath

# Get command line args
Param(
	[Parameter(Mandatory=$true)][string]$path,
	[switch]$print = $false
)

# Import libraries.
Add-Type -Path .\Libraries\PDFsharp\code\PdfSharp\bin\Debug\PdfSharp.dll

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
			Write-Host $i.Name
			$input = $PdfReader::Open($i.FullName, $PdfDocumentOpenMode::Import)
			$input.Pages | %{$output.AddPage($_)}
			# Delete processed pdfs.
			Remove-Item($i.FullName)
		} catch {
			Write-Host "Error occurred while processing file: " $i.FullName ", skipping this file."
			Write-Host "ERROR: " + $_.Exception.Message + " Failed Item: " + $_.Exception.ItemName
		}
	}
	try {
		# Create file output path
		$fileoutput = $myfolderpath + "\" + $foldername + ".pdf"
		$output.Save($fileoutput)
	} catch {
		Write-Host "Error occurred while saving file: " $fileoutput ", skipping this file."
		Write-Host "ERROR: " $_.Exception.Message " Failed Item: " $_.Exception.ItemName
	}
	
}

if ($print -eq $false) {
	# Iterate through each folder in $path alphabetically.
	Get-ChildItem $path | ?{$_.PSIsContainer -and $_.GetFiles("*.pdf").Count} | Select-Object FullName | % {
		Write-Host "FullPath: " $_.FullName
		$filepath = $_.FullName
		$filename = Split-Path $_.FullName -Leaf
		$filename = $filename + " $(Get-Date -Format MM-dd-yyyy)"
		Merge-PDF $filepath $filename
	} | Sort-Object $_.FullName

	Write-Host("Done.")
} ElseIf ($print -eq $true) {
	# Merge the PDFs in $path and print to default printer.
	$filename = Split-Path $path -Leaf
	$filename = $filename + " $(Get-Date -Format MM-dd-yyyy)"
	Merge-PDF $path $filename
	$filepath = $path + "\" + $filename + ".pdf"
	Write-Host "Printing file $($filepath)..."
	Start-Process -FilePath $filepath -Verb Print

	Write-Host("Done.")
}
