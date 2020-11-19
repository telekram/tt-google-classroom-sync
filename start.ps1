$script:config = Get-Content -Raw -Path .\config.json | ConvertFrom-Json
Import-Module .\lib\dataSources.psm1 -Force -Scope Local
Import-Module .\lib\sessionManager.psm1 -Force -Scope Local

$session = Get-ScriptSessionType

Get-ClassNames


# Invoke-Command -Session $session -ScriptBlock {
# 	$gCourses = gam print courses teacher superadmin@cheltsec.vic.edu.au | Out-String
	
# 	$props = $gCourses | ConvertFrom-Csv -Delim ','

# 	$props | Get-Member

# 	$props | ForEach-Object {
# 		$_.ID
# 		$_.ALIAS
# 		$_.NAME 
# 		$_.ROOM
# 		$_.SECTION

# 		gam update course $_.id status archived
# 	}
# 	gam create course alias 2A2D034 name "2A2D093" description 'Year2020' heading 'Studio Art' section 'Studio Art 2D Year 9 Sem 2' room 2020 status active teacher superadmin@cheltsec.vic.edu.au
	
# }

$r = Get-PSSession -ComputerName gads1.curric.cheltenham-sc.wan -UseSSL -Credential $cred
$r | Remove-PSSession
