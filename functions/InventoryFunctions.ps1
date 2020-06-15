function Get-CapaUnitHardwareInventory
{
	[CmdletBinding()]
	param
	(
        [Parameter(Mandatory = $true)]
        [string]$UnitName,
        [Parameter(Mandatory = $false)]
		[string]$Category
	)
	
	Begin
	{
		$CapaCom = New-Object -ComObject CapaInstaller.SDK
		$Capahardware = @()
	}
	Process
	{
			$Hardware = $CapaCom.GetHardwareInventoryForUnit("$UnitName", "Computer")
			$Hardwarelist = $Hardware -split "`r`n"
	
			$Hardwarelist | ForEach-Object -Process {
				$SplitLine = ($_).split('|')
				
				Try
				{
                    $CapaCustom += [pscustomobject][ordered] @{
                        Category = $SplitLine[0]
                        Entry = $SplitLine[1]
                        Value = $SplitLine[2]
                    }
				}
				Catch
				{
					Write-Warning -Message "An error occured for computer: $($SplitLine[0]) "
				}
			}
		
	}
	End
	{
		Return $Capahardware | Where-Object {$_.Category -match $Category} | Format-Table
		$CapaCom = $null
		Remove-Variable -Name CapaCom
	}
}

Function Get-CapaUnitSoftwareInventory
{
	[CmdletBinding()]
	param
	(
        [Parameter(Mandatory = $true)]
        [string]$UnitName,
        [Parameter(Mandatory = $false)]
		[string]$SoftwareName
	)
	
	Begin
	{
		$CapaCom = New-Object -ComObject CapaInstaller.SDK
		$CapaSoftware = @()
	}
	Process
	{
			$Software = $CapaCom.GetSoftwareInventoryForUnit("$UnitName", "Computer")
			$softwarelist = $Software -split "`r`n"
	
			$softwarelist | ForEach-Object -Process {
				$SplitLine = ($_).split('|')
				
				Try
				{
                    $CapaCustom += [pscustomobject][ordered] @{
                        SoftwareName = $SplitLine[0]
                        Entry = $SplitLine[1]
                        Value = $SplitLine[2]
                    }
				}
				Catch
				{
					Write-Warning -Message "An error occured for computer: $($SplitLine[0]) "
				}
			}
		
	}
	End
	{
		Return $CapaSoftware | Where-Object {$_.SoftwareName -match $SoftwareName} | Format-Table
		$CapaCom = $null
		Remove-Variable -Name CapaCom
	}
}

Function Get-CapaUnitUpdatesInventory
{

    #Note the SDK seems to have trouble pulling installdate correctly
	[CmdletBinding()]
	param
	(
        [Parameter(Mandatory = $true)]
        [string]$UnitName,
        [Parameter(Mandatory = $false)]
		[string]$UpdateName
	)
	
	Begin
	{
		$CapaCom = New-Object -ComObject CapaInstaller.SDK
		$CapaUpdates = @()
	}
	Process
	{
			$Updates = $CapaCom.GetUpdateInventoryForUnit("$UnitName", "Computer")
			$UpdatesList = $Updates -split "`r`n"
	
			$UpdatesList | ForEach-Object -Process {
				$SplitLine = ($_).split('|')
				
				Try
				{
                    $CapaUpdates += [pscustomobject][ordered] @{
                        UpdateName = $SplitLine[3]
                        Status = $SplitLine[6]
                        InstallDate = $SplitLine[1]
                    }
				}
				Catch
				{
					Write-Warning -Message "An error occured for computer: $($SplitLine[0]) "
				}
			}
		
	}
	End
	{
		Return $CapaUpdates | Where-Object {$_.UpdateName -match $UpdateName} | Format-Table
		$CapaCom = $null
		Remove-Variable -Name CapaCom
	}
}


Function Get-CapaUnitLogonHistory
{

	[CmdletBinding()]
	param
	(
        [Parameter(Mandatory = $true)]
        [string]$UnitName
	)
	
	Begin
	{
		$CapaCom = New-Object -ComObject CapaInstaller.SDK
		$CapaLogon = @()
	}
	Process
	{
			$Login = $CapaCom.GetLogonHistoryForUnit("$UnitName", "1")
			$LoginList = $Login -split "`r`n"
	
			$LoginList | ForEach-Object -Process {
				$SplitLine = ($_).split('|')
				
				Try
				{

                    if ($SplitLine[1] -match "time")
                    {
                        $SplitLine[2] = [DateTime]::FromFileTime($SplitLine[2])
                    }
                    $CapaLogon += [pscustomobject][ordered] @{
                        Category = $SplitLine[1]
                        Value = $SplitLine[2]
                    }
				}
				Catch
				{
					Write-Warning -Message "An error occured for computer: $($SplitLine[0]) "
				}
			}
		
	}
	End
	{
		Return $CapaLogon
		$CapaCom = $null
		Remove-Variable -Name CapaCom
	}
}


function Get-CapaUnitBitLocker
{
	[CmdletBinding()]
	param
	(
        [Parameter(Mandatory = $false)]
		[string]$UnitName
	)
	
	Begin
	{
		$CapaCom = New-Object -ComObject CapaInstaller.SDK
		$CapaCustom = @()
	}
	Process
	{
		Get-CapaUnit -UnitType Computer | Where-Object {$_.UnitName -match $UnitName} | ForEach-Object -Process {
			$UnitNameL = $_.UnitName
			$Custom = $CapaCom.GetCustomInventoryForUnit("$UnitNameL", "Computer")
			$CustomList = $Custom -split "`r`n"
	
			$CustomList | ForEach-Object -Process {
				$SplitLine = ($_).split('|')
				
				Try
				{
					If ($Splitline[0] -match "Bitlocker") {
						$CapaCustom += [pscustomobject][ordered] @{
							UnitName = $UnitNameL
							BitlockerKey = $SplitLine[2]
						}
					}
				}
				Catch
				{
					Write-Warning -Message "An error occured for computer: $($SplitLine[0]) "
				}
			}
		}
	}
	End
	{
		Return $CapaCustom
		$CapaCom = $null
		Remove-Variable -Name CapaCom
	}
}
