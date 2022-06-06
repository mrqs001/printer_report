$printerlist = printerlist.txt -header Value, Name, Description, ColorMode
$result = ""
$SNMP = new-object -ComObject olePrn.OleSNMP
$total = $printerlist | Measure-Object | Select-Object -expand count

$result += @"
<html>
    <head>
        <title>Printer Report</title>
        <style>* {font-family:'Trebuchet MS';}</style>
    </head>
    <body>
"@
Write-Output "Reporting on $total printers"
$x = 0

foreach ($p in $printerlist) {
    if ($p.value -notlike "-*") {

        $x = $x + 1
        $printertype = $nul
        $status = $nul
        if (!(test-connection $p.Value -Quiet -count 1)) { $result += ($p.value + " is offline<br>") }
        if (test-connection $p.value -quiet -count 1) {
            $snmp.open($p.value, "public", 2, 3000)
            $printertype = $snmp.Get(".1.3.6.1.2.1.25.3.2.1.3.1")
            Write-Output ([string]$x + ": " + [string]$p.Value + " " + $printertype)
        }
        $blacktonervolume = $snmp.get("43.11.1.1.8.1.1")
        $blackcurrentvolume = $snmp.get("43.11.1.1.9.1.1")
        [int]$blackpercentremaining = ($blackcurrentvolume / $blacktonervolume) * 100 
        if ($p.ColorMode -eq 1) {
            $cyantonervolume = $snmp.get("43.11.1.1.8.1.2")
            $cyancurrentvolume = $snmp.get("43.11.1.1.9.1.2")
            [int]$cyanpercentremaining = ($cyancurrentvolume / $cyantonervolume) * 100
            $magentatonervolume = $snmp.get("43.11.1.1.8.1.3")
            $magentacurrentvolume = $snmp.get("43.11.1.1.9.1.3")
            [int]$magentapercentremaining = ($magentacurrentvolume / $magentatonervolume) * 100
            $yellowtonervolume = $snmp.get("43.11.1.1.8.1.4")
            $yellowcurrentvolume = $snmp.get("43.11.1.1.9.1.4")
            [int]$yellowpercentremaining = ($yellowcurrentvolume / $yellowtonervolume) * 100
        }
        $statustree = $snmp.gettree("43.18.1.1.8")
        $status = $statustree | Where-Object { $_ -notlike "print*" }
        $status = $status | Where-Object { $_ -notlike "*bypass*" }
        $name = $snmp.get(".1.3.6.1.2.1.1.5.0")
        if ($name -notlike "PX*") { $name = $p.name }
             
        $result += ("<b>" + $p.description + "</b><a style='text-decoration:none;font-weight:bold;' href=http://" + $p.value + " target='_new'> " + $name + "</a> <br>" + $printertype + "<br>")
        if ($blackpercentremaining -gt 49) { $result += "<b style='font-size:110%;color:green;'>", $blackpercentremaining, "</b>% Preto<br>" }
        if (($blackpercentremaining -gt 24) -and ($blackpercentremaining -le 49)) { $result += "<b style='font-size:110%;color:#40BB30;'>", $blackpercentremaining, "</b>% Preto<br>" }
        if (($blackpercentremaining -gt 10) -and ($blackpercentremaining -le 24)) { $result += "<b style='font-size:110%;color:orange;'>", $blackpercentremaining, "</b>% Preto<br>" }
        if (($blackpercentremaining -ge 0) -and ($blackpercentremaining -le 10)) { $result += "<b style='font-size:110%;color:red;'>", $blackpercentremaining, "</b>% Preto<br>" }
        if ($p.ColorMode -eq 1) {
            if ($cyanpercentremaining -gt 49) { $result += "<b style='font-size:110%;color:green;'>", $cyanpercentremaining, "</b>% Ciano<br>" }
            if (($cyanpercentremaining -gt 24) -and ($cyanpercentremaining -le 49)) { $result += "<b style='font-size:110%;color:#40BB30;'>", $cyanpercentremaining, "</b>% Ciano<br>" }
            if (($cyanpercentremaining -gt 10) -and ($cyanpercentremaining -le 24)) { $result += "<b style='font-size:110%;color:orange;'>", $cyanpercentremaining, "</b>% Ciano<br>" }
            if (($cyanpercentremaining -ge 0) -and ($cyanpercentremaining -le 10)) { $result += "<b style='font-size:110%;color:red;'>", $cyanpercentremaining, "</b>% Ciano<br>" }
            if ($magentapercentremaining -gt 49) { $result += "<b style='font-size:110%;color:green;'>", $magentapercentremaining, "</b>% Magenta<br>" }
            if (($magentapercentremaining -gt 24) -and ($magentapercentremaining -le 49)) { $result += "<b style='font-size:110%;color:#40BB30;'>", $magentapercentremaining, "</b>% Magenta<br>" }
            if (($magentapercentremaining -gt 10) -and ($magentapercentremaining -le 24)) { $result += "<b style='font-size:110%;color:orange;'>", $magentapercentremaining, "</b>% Magenta<br>" }
            if (($magentapercentremaining -ge 0) -and ($magentapercentremaining -le 10)) { $result += "<b style='font-size:110%;color:red;'>", $magentapercentremaining, "</b>% Magenta<br>" }
            if ($yellowpercentremaining -gt 49) { $result += "<b style='font-size:110%;color:green;'>", $yellowpercentremaining, "</b>% Amarelo<br>" }
            if (($yellowpercentremaining -gt 24) -and ($yellowpercentremaining -le 49)) { $result += "<b style='font-size:110%;color:#40BB30;'>", $yellowpercentremaining, "</b>% Amarelo<br>" }
            if (($yellowpercentremaining -gt 10) -and ($yellowpercentremaining -le 24)) { $result += "<b style='font-size:110%;color:orange;'>", $yellowpercentremaining, "</b>% Amarelo<br>" }
            if (($yellowpercentremaining -ge 0) -and ($yellowpercentremaining -le 10)) { $result += "<b style='font-size:110%;color:red;'>", $yellowpercentremaining, "</b>% Amarelo<br>" }
        }
        if ($status.length -gt 0) { $result += ($status + "<br><br>") }else { $result += "Operacional<br><br>" }
    }
}

$result += "</body></html>"

$EmailParams = @{
    To         = ""
    From       = ""
    Subject    = "Relatório Diário Impressoras"
    Body       = $result
    SMTPServer = ''
    Encoding   = [System.Text.Encoding]::GetEncoding('iso-8859-1')
}
#Send E-mail from PowerShell script
Send-MailMessage @EmailParams -BodyAsHtml
