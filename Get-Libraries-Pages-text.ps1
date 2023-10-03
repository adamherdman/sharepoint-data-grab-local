# Load required assemblies
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Open-XML-SDK.2.5.0\lib\net45\DocumentFormat.OpenXml.dll" 

###############################################################################
# PPTX: Extract text from PowerPoint slides
function ExtractTextFromPptx($pptxPath) {
    $presentation = [DocumentFormat.OpenXml.Packaging.PresentationDocument]::Open($pptxPath, $false)
    $textBuilder = New-Object System.Text.StringBuilder

    foreach ($slide in $presentation.PresentationPart.SlideParts) {
        $textBuilder.AppendLine($slide.Slide.InnerText)
    }

    $presentation.Close()
    return $textBuilder.ToString()
}
###############################################################################
# PDF: Extract text using pdftotext tool
function ExtractTextFromPdf($pdfPath) {
    $output = "$pdfPath.temp.txt"
    & "$PSScriptRoot\xpdf-tools-win-4.04\bin32\pdftotext.exe" $pdfPath $output

    $content = Get-Content $output
    Remove-Item $output -Force
    return $content
}
###############################################################################
# DOCX: Extract text from Word document
function ExtractTextFromDocx($docxPath) {
    $document = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Open($docxPath, $false)
    $text = $document.MainDocumentPart.Document.Body.InnerText
    $document.Close()
    return $text
}
###############################################################################
# ASPX: Decode and clean HTML content
function ExtractAndCleanTextFromAspx($aspxPath) {
    $rawContent = Get-Content -Path $aspxPath
    $decodedContent = [System.Web.HttpUtility]::HtmlDecode($rawContent)
    $cleanedContent = $decodedContent -replace '<[^>]+>', ''
    return $cleanedContent
}
###############################################################################


# SharePoint Online credentials and site setup
# Import the credentials from the credentials file
. .\credentials.ps1
# Use the credentials for SharePoint Online
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password)
$siteUrl = "https://nhs.sharepoint.com/sites/msteams_4fc4c8"

# Connect to SharePoint Online
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$ctx.Credentials = $credentials

# Load document libraries (ignoring hidden ones)
$lists = $ctx.Web.Lists
$ctx.Load($lists)
$ctx.ExecuteQuery()

# Output directory setup
$outputDir = "$PSScriptRoot\DataOutput"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}

# Desired file extensions to process
$desiredExtensions = @(".pdf", ".aspx", ".docx", ".pptx")
$baseUrl = "https://nhs.sharepoint.com"

# Iterate over each document in libraries and process based on extension
foreach ($list in $lists) {
    if (-not $list.Hidden) {
        $items = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
        $ctx.Load($items)
        $ctx.ExecuteQuery()

        foreach ($item in $items) {
            $url = $item["FileRef"]
            $extension = [System.IO.Path]::GetExtension($url)

            if ($desiredExtensions -contains $extension) {
                $file = $ctx.Web.GetFileByServerRelativeUrl($url)
                $ctx.Load($file)
                $ctx.ExecuteQuery()

                $binaryStreamResult = $file.OpenBinaryStream()
                $ctx.ExecuteQuery()

                $tempFile = "$outputDir\$([System.IO.Path]::GetFileName($url))"
                $fileStream = [System.IO.File]::Create($tempFile)
                $binaryStreamResult.Value.CopyTo($fileStream)
                $fileStream.Close()

                $content = switch ($extension) {
                    ".pdf"  { ExtractTextFromPdf($tempFile) }
                    ".docx" { ExtractTextFromDocx($tempFile) }
                    ".aspx" { ExtractAndCleanTextFromAspx($tempFile) }
                    ".pptx" { ExtractTextFromPptx($tempFile) }
                }

                $fullUrl = "$baseUrl$($url -replace ' ', '%20')"
                $content = "$fullUrl`r`n$content"
                $outputFile = Join-Path $outputDir ([System.IO.Path]::GetFileNameWithoutExtension($url) + ".txt")
                
                Set-Content -Path $outputFile -Value $content
                Remove-Item -Path $tempFile
            }
        }
    }
}
