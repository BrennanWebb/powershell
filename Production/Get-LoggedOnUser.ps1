Function Get-LoggedInUser {
    <#
    .SYNOPSIS
        This will check the specified machine to see all users who are logged on.
        For updated help and examples refer to -Online version.
     
    .NOTES
        Name: Get-LoggedInUser
        Author: theSysadminChannel
        Version: 2.0
        DateCreated: 2020-Apr-01
     
     
    .LINK
        https://thesysadminchannel.com/get-logged-in-users-using-powershell/ -
        For updated help and examples refer to -Online version.
     
     
     
    .PARAMETER ComputerName
        Specify a computername to see which users are logged into it.  If no computers are specified, it will default to the local computer.
     
    .PARAMETER UserName
        If the specified username is found logged into a machine, it will display it in the output.
     
    PARAMETER Logoff
        Logoff the users from the computers in your query. It is recommended to run without the logoff switch to view the results.
     
    .EXAMPLE
        Get-LoggedInUser -ComputerName Server01
     
        Display all the users that are logged in server01
     
    .EXAMPLE
        Get-LoggedInUser -ComputerName Server01, Server02 -UserName jsmith
     
        Display if the user, jsmith, is logged into server01 and/or server02
     
    .EXAMPLE
        Get-LoggedInUser -ComputerName $ComputerList -Logoff
     
        Logoff all the users that are logged into the computers in the ComputerList array
     
    .EXAMPLE
        Get-LoggedInUser -ComputerName $ComputerList -SamAccountName jsmith -Logoff
     
        If jsmith is logged into a computer in the $ComputerList array, it will log them out.
     
    #>
     
        [CmdletBinding(DefaultParameterSetName="Default")]
            param(
                [Parameter(
                    Mandatory = $false,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true,
                    Position=0
                )]
                [string[]] $ComputerName = $env:COMPUTERNAME,
     
     
                [Parameter(
                    Mandatory = $false
                )]
                [Alias("SamAccountName")]
                [string]   $UserName,
     
     
                [Parameter(
                    Mandatory = $false
                )]
                [switch]   $Logoff
            )
     
        BEGIN {}
     
        PROCESS {
            Foreach ($Computer in $ComputerName) {
                try {
                    $ExplorerProcess = Get-WmiObject Win32_Process -Filter "Name = 'explorer.exe'" -ComputerName $Computer -ErrorAction Stop
                    $Computer = $Computer.ToUpper()
                    if ($ExplorerProcess) {
                        $UserList = $ExplorerProcess.GetOwner() | Where-Object {$_.ReturnValue -eq 0} | select -ExpandProperty User
                        if ($ExplorerProcess.GetOwner() | Where-Object {$_.ReturnValue -ne 0}) {
                            Write-Warning "Other users are logged in to $($Computer) but couldn't pull the details.  Consider running as an administrator."
                        }
     
                        if (-not $PSBoundParameters.ContainsKey("UserName") -and -not $PSBoundParameters.ContainsKey("Logoff")) {
                            foreach ($User in $UserList) {
                                $User = $User.ToLower()
                                $Session = (query session $User /Server:$Computer | Select-String -Pattern $User -ErrorAction Stop).ToString().Trim()
                                $Session = $Session -replace '\s+', ' '
                                $Session = $Session -replace '>', ''
     
                                if ($Session.Split(' ')[2] -cne "Disc") {
                                    [PSCustomObject]@{
                                        ComputerName = $Computer
                                        UserName     = $User.Replace('{}','')
                                        SessionID    = $Session.Split(' ')[2]
                                        SessionState = $Session.Split(' ')[3]
                                    }
                                  } else {
                                    [PSCustomObject]@{
                                        ComputerName = $Computer
                                        UserName     = $User.Replace('{}','')
                                        SessionID    = $Session.Split(' ')[1]
                                        SessionState = 'Disconnected'
                                    }
                                }
                            }
                        } #End Default PSBoundParameter Block
     
                        if ($PSBoundParameters.ContainsKey("UserName") -and -not $PSBoundParameters.ContainsKey("Logoff")) {
                            foreach ($User in $UserList) {
                                if ($User -eq $UserName) {
                                    $User = $User.ToLower()
                                    $Session = (query session $User /Server:$Computer | Select-String -Pattern $User -ErrorAction Stop).ToString().Trim()
                                    $Session = $Session -replace '\s+', ' '
                                    $Session = $Session -replace '>', ''
     
                                    if ($Session.Split(' ')[2] -cne "Disc") {
                                        [PSCustomObject]@{
                                            ComputerName = $Computer
                                            UserName     = $User.Replace('{}','')
                                            SessionID    = $Session.Split(' ')[2]
                                            SessionState = $Session.Split(' ')[3]
                                        }
                                      } else {
                                        [PSCustomObject]@{
                                            ComputerName = $Computer
                                            UserName     = $User.Replace('{}','')
                                            SessionID    = $Session.Split(' ')[1]
                                            SessionState = 'Disconnected'
                                        }
                                    }
                                }
                            }
                        } #End UserName PSBoundParameter Block
     
                        if ($PSBoundParameters.ContainsKey("Logoff") -and -not $PSBoundParameters.ContainsKey("UserName")) {
                            foreach ($User in $UserList) {
                                $User = $User.ToLower()
                                $Session = (query session $User /Server:$Computer | Select-String -Pattern $User -ErrorAction Stop).ToString().Trim()
                                $Session = $Session -replace '\s+', ' '
                                $Session = $Session -replace '>', ''
     
                                if ($Session.Split(' ')[2] -cne "Disc") {
                                    LogOff.exe /server:$Computer ($Session.Split(' ')[2])
     
                                    [PSCustomObject]@{
                                        ComputerName = $Computer
                                        UserName     = $User.Replace('{}','')
                                        SessionID    = $Session.Split(' ')[2]
                                        SessionState = 'LoggingOff'
                                    }
                                } else {
                                    LogOff.exe /server:$Computer ($Session.Split(' ')[1])
     
                                    [PSCustomObject]@{
                                        ComputerName = $Computer
                                        UserName     = $User.Replace('{}','')
                                        SessionID    = $Session.Split(' ')[1]
                                        SessionState = 'LoggingOff'
                                    }
                                }
                            }
                        } #End Logoff PSBoundParameter Block
     
                        if ($PSBoundParameters.ContainsKey("Logoff") -and $PSBoundParameters.ContainsKey("UserName")) {
                            foreach ($User in $UserList) {
                                if ($User -eq $UserName) {
                                    $User = $User.ToLower()
                                    $Session = (query session $User /Server:$Computer | Select-String -Pattern $User -ErrorAction Stop).ToString().Trim()
                                    $Session = $Session -replace '\s+', ' '
                                    $Session = $Session -replace '>', ''
     
                                    if ($Session.Split(' ')[2] -cne "Disc") {
                                        LogOff.exe /server:$Computer ($Session.Split(' ')[2])
     
                                        [PSCustomObject]@{
                                            ComputerName = $Computer
                                            UserName     = $User.Replace('{}','')
                                            SessionID    = $Session.Split(' ')[2]
                                            SessionState = 'LoggingOff'
                                        }
                                    } else {
                                        LogOff.exe /server:$Computer ($Session.Split(' ')[1])
     
                                        [PSCustomObject]@{
                                            ComputerName = $Computer
                                            UserName     = $User.Replace('{}','')
                                            SessionID    = $Session.Split(' ')[1]
                                            SessionState = 'LoggingOff'
                                        }
                                    }
                                }
                            }
                        } #End Logoff PSBoundParameter Block
                    }
     
                } catch {
                    Write-Error "$($Computer.ToUpper()) - $($_.Exception.Message)"
     
                }
            }
        }
     
        END {}
    };


    Get-LoggedInUser;