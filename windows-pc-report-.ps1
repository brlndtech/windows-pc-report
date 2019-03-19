######################################
######################################

# Windows PC Report | Brlndtech © 

######################################
######################################
#### HTML Output Formatting #######

$a = "<style>"
$a = $a + "BODY{background-color:#f2f2f2 ;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:LightCyan}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:AliceBlue}"
$a = $a + "</style>"
$a = $a + "
<!DOCTYPE html> 
<html>
<head>
<title>Windows Pc Report</title>
</head>
<body>
<h1></h1>
<p></p>
</body>
</html>"

################################################################################################

###### Global variables ####

$vUserName = (Get-Item env:\username).Value 			## This will get username using environment variable
$vComputerName_report = (Get-Item env:\Computername).Value     ## this is computer name using environment variable
$filepath = (Get-ChildItem env:\userprofile).value	## this is user profile  using environment variable



ConvertTo-Html -Title "System Information for $vComputerName_report" -Body " <br> <center> <div style=""width:1200px;height:980px;border:2px solid #999999;""> <br> <!-- logo de votre organisation  --><img src=""https://www.lyon-metropole.cci.fr/plugins/CciWebPlugin/images/logo_CCI-Lyon-Metropole-Saint-Etienne-Roanne.png"" width=200 align=""top"" /> <center> <h1> Status Report for : $vComputerName_report </h1>" >  "$filepath\Desktop\$vComputerName_report.html" 

################################################
#  Hardware Information
#################################################

# ConvertTo-Html -Body "<H1>HARDWARE INFORMATION </H1>" >> "$filepath\Desktop\$vComputerName_report.html"
						  

################################################
#  OS Information
#################################################

ConvertTo-Html -Body "<H1>OS INFORMATION </H1>" >> "$filepath\$name.html" 

get-WmiObject win32_operatingsystem -ComputerName $vComputerName_report | select Caption,InstallDate,OSArchitecture,Version,BootDevice `
                                          | ConvertTo-html -Body "<H3> Operating System Information</H3>" >>  "$filepath\Desktop\$vComputerName_report.html"
################################################
#  Network organization & more 
#################################################
Get-WMIObject -class Win32_ComputerSystem | select Domain, PrimaryOwnerName, Manufacturer, Model | ConvertTo-html -Body "<H3> Network Organization & more  </H3>" >>  "$filepath\Desktop\$vComputerName_report.html"

################################################
#  Physical and logical disk(s)
#################################################

# Get-WmiObject win32_DiskDrive -ComputerName $vComputerName_report | Select Model,SerialNumber,Description,MediaType,FirmwareRevision |ConvertTo-html -Body "<h3> Physical DISK Drives </h3>" >>  "$filepath\Desktop\$vComputerName_report.html"							  
Get-WmiObject win32_logicalDisk -ComputerName $vComputerName_report | select DeviceID,VolumeName,ProviderName,@{Expression={$_.Size /1Gb -as [int]};Label="Total Size(GB)"},@{Expression={$_.Freespace / 1Gb -as [int]};Label="Free Size (GB)"} `
                                         | ConvertTo-html -Body "<h3> Logical DISK Drives </h3>" >>  "$filepath\Desktop\$vComputerName_report.html"

################################################
#  Networks IP /x /DHCP 
#################################################									 
(Get-NetIPConfiguration).IPv4Address | Select  IPAddress, PrefixLength, InterfaceAlias, SuffixOrigin  | ConvertTo-html -Body "<h3> Network(s) </h3>" >>  "$filepath\Desktop\$vComputerName_report.html" 

################################################
#  Routing informations  
################################################# 
Get-WmiObject -Class Win32_IP4RouteTable |
  where { $_.destination -eq '0.0.0.0' -and $_.mask -eq '0.0.0.0'} |
  Sort-Object metric1 | select nexthop, metric1, interfaceindex, Age | ConvertTo-html -Body "<h3> Routing (to access to the internet) </h3>" >>  "$filepath\Desktop\$vComputerName_report.html"
  
################################################
#  CPU information 
################################################   
  	 									 
Get-WmiObject win32_processor | select DeviceID, Name, MaxClockSpeed | ConvertTo-html -Body "<h3> CPU Information </h3>" >>  "$filepath\Desktop\$vComputerName_report.html"

################################################
#  footer   
#################################################

$Report = " <br> <br> <br> <strong> <i> <h4>The report was generated On  $(get-date) by <u> $((Get-Item env:\username).Value) </u> (post deployment) 
</h4> </i> </strong> "
$Report  >> "$filepath\Desktop\$vComputerName_report.html" 
$report2 = " <h4> All rights reserved - brlndtech © |  More Info ? <a href =""https://github.com/brlndtech/windows-pc-report"" target=""_blank""> check my github </a> <br> </strong> </div> </center> "
$Report2  >> "$filepath\Desktop\$vComputerName_report.html" 
$report2 = " </div>"
$vComputerName_report | ConvertTo-html  -Head $a -Body "" >>  "$filepath\Desktop\$vComputerName_report.html"

##################################################
# automatically opening the page 
##################################################
invoke-Expression "$filepath\Desktop\$vComputerName_report.html"  
#################### END OF THE SCRIPT  ####################################
#################### END OF THE  SCRIPT ####################################