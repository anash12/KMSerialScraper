#Title: KMSerialScraper
#Author: Drew Nash
#Date: 11/8/2018
#Version: 1.0
#
#Description: Wrote this script in a few minutes to help with a project at work. Scrapes Konica Minolta printer web interface for it's serial number and outputs it to a CSV file. this works on the most recent Konica Minolta firmware.
#             The script requires a list of IP addresses within a CSV file to work.
#


#import list of printers/IPs in CSV file
$CSV = @()
$CSV = Import-Csv C:\Temp\printers.csv

ForEach($printer in $CSV){
    #create IE instance
    $IE = New-Object -ComObject InternetExplorer.Application
    $IE.Visible = $true
    $IE.Navigate2($printer.IP)
    while($IE.ReadyState -ne 4){Start-Sleep -Milliseconds 100}
    #Bypass SSL cert error if present
    if($IE.Document.url -like "*invalidcert*"){
        $sslbypass=$ie.Document.getElementsByTagName("a") | where-object {$_.id -eq "overridelink"};
        $sslbypass.click();
        "sleep for 15 seconds while final page loads";
        start-sleep -s 15;
    }
    Write-Host "Looking up serial for "$printer.IP"..."
    while($ie.Busy) { Start-Sleep -Milliseconds 100 }
    #get the HTML element containing the serial number
    $serialNumber = $IE.Document.IHTMLDocument3_getElementsByTagName("td")[24].innerhtml
    Write-Host $serialNumber
    $printer.Serial = $serialNumber
    $IE.Quit()
}

$CSV | Export-Csv C:\Temp\printersWithSerials.csv
