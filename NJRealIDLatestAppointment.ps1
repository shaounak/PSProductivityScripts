$url = "https://telegov.njportal.com/njmvc/AppointmentWizard/12"
$ie = New-Object -com internetexplorer.application
$ie.visible = $true
$ie.navigate($url)
while ($ie.ReadyState -ne 4) { start-sleep -m 2000 }

$datebtns = $($ie.Document.IHTMLDocument3_getElementsByTagName('a') | where-object {$_.id -match 'datebtn' })
ForEach ($datebtn in $datebtns) {
        $datebtn.click()
        Start-Sleep -Seconds 1
}
Start-Sleep -Seconds 1
$ie.Document.IHTMLDocument3_getElementsByTagName("span") `
| Where-Object {$_.id -match 'dateText'} `
| Where-Object {$_.innerText -notcontains 'No Appointments Available'} `
| Out-GridView
