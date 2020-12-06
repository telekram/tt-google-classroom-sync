$script:config = Get-Content -Raw -Path .\config.json | ConvertFrom-Json

function Get-ScriptPSSession () {
  $isRemoteSession = [System.Convert]::ToBoolean($script:config.remoteSession)

  $sessionUser = $script:config.username
  $sessionPWord = ConvertTo-SecureString -String $script:config.password -AsPlainText -Force

  $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sessionUser, $sessionPWord
  

  if ($isRemoteSession) {
    #$session = New-PSSession -ComputerName gads1.curric.cheltenham-sc.wan -UseSSL -Credential $cred
    $session = New-PSSession -ComputerName gads1.curric.cheltenham-sc.wan -Credential $cred
  } else {
    $session = New-PSSession 
  }

  return $session
}

function Clear-ScriptPSSession {
  $isRemoteSession = [System.Convert]::ToBoolean($script:config.remoteSession)

  $sessionUser = $script:config.username
  $sessionPWord = ConvertTo-SecureString -String $script:config.password -AsPlainText -Force

  $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sessionUser, $sessionPWord
  if ($isRemoteSession) {
    #$r = Get-PSSession -ComputerName gads1.curric.cheltenham-sc.wan -UseSSL -Credential $cred
    $r = Get-PSSession -ComputerName gads1.curric.cheltenham-sc.wan -Credential $cred
    $r | Remove-PSSession
  }
}