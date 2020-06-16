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

			$CapaLogon | ForEach-Object -Process {
				if ($_.Category -match "time" -or $_.Category -match "Inventory Collected")
				{
					$_.Value = (Get-Date 01.01.1970)+([System.TimeSpan]::fromseconds($_.Value))
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


Function Get-CapaUserInventory
{

	[CmdletBinding()]
	param
	(
        [Parameter(Mandatory = $true)]
        [string]$UserName
	)
	
	Begin
	{
		$CapaCom = New-Object -ComObject CapaInstaller.SDK
		$CapaUser = @()
	}
	Process
	{
			$User = $CapaCom.GetUserInventory("$UserName")
			$UserList = $User -split "`r`n"
	
			$UserList | ForEach-Object -Process {
				$SplitLine = ($_).split('|')

				$SplitLine | ForEach-Object -Process {
					
				}
				
				

				Try
				{
                    $CapaUser += [pscustomobject][ordered] @{
                        Entry = $SplitLine[1]
                        Value = $SplitLine[2]
					}
				}
				Catch
				{
					Write-Warning -Message "An error occured for computer: $($SplitLine[0]) "
				}
			}

			$CapaUser | ForEach-Object -Process {
				if ($_.Entry -match "Password last changed" -or $_.entry -match "Last failed login time" -or $_.Entry -match "Account expire date" -or $_.Entry -match "Inventory collected" -and $_.Value -ne "")
				{
					$_.Value = (Get-Date 01.01.1970)+([System.TimeSpan]::fromseconds($_.Value))
				}
			}
		
	}
	End
	{
		Return $CapaUser
		$CapaCom = $null
		Remove-Variable -Name CapaCom
	}

}
