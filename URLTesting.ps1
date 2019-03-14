[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Folder = "C:\Users\JT\Desktop\Trident Nov 9\Course resources and files\Course folders to test links"
#$$Folder = "C:\dnload\scripts\powershell\Test"
$Documenttype = "docx"
$WordFiles = get-childitem -path $folder -Recurse -include *.$Documenttype
$filename =  get-date -format MMddyyyy_hhMM
$results = @()

foreach ($File in $WordFiles.fullname) 
{
    $word = New-Object -ComObject word.application
    $document = $word.documents.open($File)
    $results += "  "
    $results += $File
    $hyperlinks = @($document.Hyperlinks) 
    $hyperlinks = $hyperlinks.addressold

    
    foreach ($item in $hyperlinks) 
    {
        $HTTP_Request = [System.Net.WebRequest]::Create($item)
        Write-Host "Testing $item"
        # Get a response from the site.  
        try{$HTTP_Response = $HTTP_Request.GetResponse()}
          catch [System.Net.WebException] {
                Write-host 'Bad Site' `n -foregroundcolor Red
                $results+=$item
                continue
            }
      
        $HTTP_Response = $HTTP_Request.GetResponse()
        # Get the HTTP code as an integer.
        $HTTP_Status = $null    
        $HTTP_Status = [int]$HTTP_Response.StatusCode
            If ($HTTP_Status -eq 200) {
            Write-Host "Site is OK!" `n -foregroundcolor Green
            }
            else {
                write-host 'Bad Site  -- Please review'
            }
            
        # Clean up the http request by closing it.
        $HTTP_Response.Close()
    }
#quit WinWord
    $word.quit()
}

#print out Results
$results >  $Folder\$filename.csv
