<#
You will need have the "ImportExcel" Module installed for this to properly run. 
You can get it here:
https://www.powershellgallery.com/packages/ImportExcel/7.4.1
To install it run: 
Install-Module -Name ImportExcel -RequiredVersion 7.4.1
Import-Module -Name ImportExcel
This will pull the basic environment from the Cisco Call Manager. 

It will place the Excel spreadsheet it in the location you enter when prompted. 

Note: RIS only supports 1000 devices per query, adjust as needed. Line 61 -<soap:MaxReturnedDevices>1000</soap:MaxReturnedDevices>
https://developer.cisco.com/docs/sxml/#!risport70-api-reference
#>


Function Write-DataToExcel
    {
        param ($filelocation, $details, $tabname)

        
        $excelpackage = Open-ExcelPackage -Path $filelocation 
        $ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName $tabname
        $details | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter
        Clv details 

    }


Function Get-CUCM_Environment
{
    param ($filelocation)
    $Details = @()

    # Enter CUCM Server
     $cucmServer = Read-Host "Enter CUCM Server IP"
     $ver = Read-Host "Enter CUCM Verison (ex 10.5, 11.0, 12.5, 14.0):"
    # Max Lines on  Device
    $maxlines = 20

    # Enter AXL User/Pass
    $user = Read-Host "AXL Username"
    $password = Read-Host "AXL Password"
   

    # Set Credentials
    $pass = ConvertTo-SecureString $password -AsPlainText -Force
    $cred = New-Object Management.Automation.PSCredential ($user, $pass)
  
    ############################################################################################################################################
    # Get RIS Data
    ############################################################################################################################################
    $request = @"
    <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:soap="http://schemas.cisco.com/ast/soap">
        <soapenv:Header/>
        <soapenv:Body>
            <soap:selectCmDevice>
                <soap:StateInfo></soap:StateInfo>

                <soap:CmSelectionCriteria>
                <soap:MaxReturnedDevices>1000</soap:MaxReturnedDevices>
                <soap:DeviceClass>Any</soap:DeviceClass>
                <soap:Model>255</soap:Model>
                <soap:Status>Registered</soap:Status>
                <soap:NodeName></soap:NodeName>
                <soap:SelectBy>DirNumber</soap:SelectBy>
                <soap:SelectItems>
                <!--Zero or more repetitions:-->
                <soap:item>
									<soap:Item>*</soap:Item>
								</soap:item>
                </soap:SelectItems>
                <soap:Protocol>Any</soap:Protocol>
                <soap:DownloadStatus>Any</soap:DownloadStatus>
                </soap:CmSelectionCriteria>
            </soap:selectCmDevice>
        </soapenv:Body>
    </soapenv:Envelope>
"@

    [System.Net.ServicePointManager]::Expect100Continue = $false

    Write-Host "Connecting to CUCM Server https://$cucmServer"

    try {
        $result = Invoke-RestMethod -Method Post -Uri "https://$cucmServer`:8443/realtimeservice2/services/RISService70?wsdl" -Headers @{'Content-Type'='text/xml';'SOAPAction'='http://schemas.cisco.com/ast/soap/action/#RisPort#SelectCmDevice'} -Body $request -Credential $cred -SkipCertificateCheck
    } catch {
        Write-Host "ERROR"
        $_ | Select-Object -ExpandProperty ErrorDetails | Select-Object -ExpandProperty Message
        exit
    }


    Write-Host "------------------------------------------------------------------"
    foreach($x in $result.Envelope.Body.selectCmDeviceResponse.selectCmDeviceReturn.SelectCmDeviceResult.CmNodes.ChildNodes){
        Write-Host "----------Parsing " $x.Name
        $cname = $x.Name
        foreach($y in $x.CmDevices.ChildNodes){
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "CUCMNode" -Value $cname   
            $detail | add-Member -MemberType NoteProperty -Name "Name" -Value $y.Name   
            #$detail | add-Member -MemberType NoteProperty -Name "DirNumber" -Value $y.DirNumber
            $detail | add-Member -MemberType NoteProperty -Name "DeviceClass" -Value $y.DeviceClass
            $detail | add-Member -MemberType NoteProperty -Name "Model" -Value $y.Model
            $detail | add-Member -MemberType NoteProperty -Name "Product" -Value $y.Product
            $detail | add-Member -MemberType NoteProperty -Name "Status" -Value $y.Status
            $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $y.Description
            $detail | add-Member -MemberType NoteProperty -Name "Protocol" -Value $y.Protocol
            $detail | add-Member -MemberType NoteProperty -Name "NumOfLines" -Value $y.NumOfLines
            $detail | add-Member -MemberType NoteProperty -Name "ActiveLoadID" -Value $y.ActiveLoadID
            $x = 0
            foreach($ip in $y.IPAddress.ChildNodes){
                $x++    
                $detail | add-Member -MemberType NoteProperty -Name "IP_$x" -Value $ip.IP
        
            }
            $x = 1
            foreach($line in $y.LinesStatus.ChildNodes){
#                Write-Host $x " - " $line.DirectoryNumber
                $detail | add-Member -MemberType NoteProperty -Name Line$x -Value $line.DirectoryNumber
                $x++
            }
            while($x -le $maxlines){
                $detail | add-Member -MemberType NoteProperty -Name Line$x -Value ""
                $x++
            }

            $Details += $detail        
        }
    }
    $Details|Export-Excel -Path $filelocation -WorksheetName "RISData" -AutoSize -AutoFilter -Show -FreezeTopRow 
    clv details
}



Write-Host "This is will create an Excel Spreadsheet.  Make sure to enter the file name with .xlsx"
Import-Module ImportExcel
$filelocation = Read-Host "Enter Location/filename to store output (i.e c:\scripts\test.xlsx)"


# Determine if ImportExcel module is loaded
$XLmodule = Get-Module -Name importexcel


if ($XLmodule )
    {
        Get-CUCM_Environment $filelocation
        Write-Host "Done"
    }
    Else {Write-Host "ImportExcel module is not loaded"}
