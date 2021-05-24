$script:config = Get-Content -Raw -Path .\config.json | ConvertFrom-Json

function Get-ScriptPSSession () {
  $isRemoteSession = $script:config.isRemoteSession

  $sessionUser = $script:config.username
  $sessionPWord = ConvertTo-SecureString -String $script:config.password -AsPlainText -Force

  $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sessionUser, $sessionPWord
  

  if ($isRemoteSession) {
    $session = New-PSSession -ComputerName gads1.curric.cheltenham-sc.wan -UseSSL -Credential $cred
  } else {
    $session = New-PSSession
  }

  return $session
}

function Clear-ScriptPSSession ($session) {
  $session | Remove-PSSession
}