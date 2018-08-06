class Device
{
    # Optionally, add attributes to prevent invalid values
    [ValidateNotNullOrEmpty()][string]$SANType
    [ValidateNotNullOrEmpty()][string]$DeviceName
    [ValidateNotNullOrEmpty()][string]$HBA
    [ValidateNotNullOrEmpty()][string]$HBAPort
    [ValidateNotNullOrEmpty()][string]$Fabric
    [ValidateNotNullOrEmpty()][string]$WWN_DN
    [ValidateNotNullOrEmpty()][string]$WWN_DP
    
    
    # optionally, have a constructor to 
    # force properties to be set:
    Device($SANType, $DeviceName, $HBA, $HBAPort, $Fabric, $WWN_DN, $WWN_DP) {
       $this.SANType = $SANType
       $this.DeviceName = $DeviceName
       $this.HBA = $HBA
       $this.HBAPort = $HBAPort
       $this.Fabric = $Fabric
       $this.WWN_DN = $WWN_DN
       $this.WWN_DP = $WWN_DP
    }

    [string] GetCreateAlias()
    {
        if ($this.SANType -eq "I") {
            return "alicreate `"{0}_H{1}_P{2}`", `"{3}; {4}`"" -f $this.DeviceName, $this.HBA, $this.HBAPort, $this.WWN_DN, $this.WWN_DP
        } elseif ($this.SANType -eq "T") {
            return "alicreate `"{0}_C{1}_P{2}`", `"{3}; {4}`"" -f $this.DeviceName, $this.HBA, $this.HBAPort, $this.WWN_DN, $this.WWN_DP            
        } else {
            return "None"
        }
    }
    [string] GetAliasName()
    {
        if ($this.SANType -eq "I") {
            return "{0}_H{1}_P{2}" -f $this.DeviceName, $this.HBA, $this.HBAPort
        } elseif ($this.SANType -eq "T") {
            return "{0}_C{1}_P{2}" -f $this.DeviceName, $this.HBA, $this.HBAPort
        } else {
            return "None"
        }
    }
}


class ZoneName
{
    # Optionally, add attributes to prevent invalid values
    [string]$LeftName
    [string]$RightName
    
    # optionally, have a constructor to 
    # force properties to be set:
    ZoneName($LeftName, $RightName)  {
       $this.LeftName = $LeftName
       $this.RightName = $RightName
    }
}
$zones = @()

$initiators = @()
$targets = @()
$fabrics = @{}

$excel = New-Object -ComObject excel.Application
$wb = $excel.workbooks.open("C:\Users\N.Artykaly\Documents\git\san-zoning\Zoning.xlsx")

for ($i=1; $i -le $wb.sheets.count; $i++){
    $sh=$wb.Sheets.Item($i)
    $sh | Select-Object -Property Name

    $Lines = $sh.UsedRange.Rows.Count

    for ($i = 2; $i -le $Lines; $i++) {
        $leftName  = $sh.Rows.Item($i).columns.Item(1).Text
        $rightName = $sh.Rows.Item($i).columns.Item(2).Text

        $ZoneNameVar = [ZoneName]::New($leftName, $rightName)

        $zones += $ZoneNameVar
    }
  
}
$excel.Workbooks.Close()

$wb = $excel.workbooks.open("C:\Users\N.Artykaly\Documents\git\san-zoning\Initiators.xlsx")

ForEach ($zoneName in $zones) {
    for ($i=1; $i -le $wb.sheets.count; $i++){
        $sh=$wb.Sheets.Item($i)
    
        if ($sh.Name -eq $zoneName.LeftName) {
            $Lines = $sh.UsedRange.Rows.Count
    
            for ($j = 2; $j -le $Lines; $j++) {
                $DeviceName = $sh.Rows.Item($j).columns.Item(1).Text
                $HBA        = $sh.Rows.Item($j).columns.Item(2).Text
                $HBAPort    = $sh.Rows.Item($j).columns.Item(3).Text
                $Fabric     = $sh.Rows.Item($j).columns.Item(4).Text
                $WWN_DN     = $sh.Rows.Item($j).columns.Item(5).Text
                $WWN_DP     = $sh.Rows.Item($j).columns.Item(6).Text
                
                if (-not ([string]::IsNullOrEmpty($Fabric))) {
                    if (! $fabrics.ContainsKey($Fabric)) {
                        $fabrics.Add($Fabric, $Fabric)
                    }

                    $device = [Device]::New("I", $DeviceName, $HBA, $HBAPort, $Fabric, $WWN_DN, $WWN_DP)
    
                    $initiators += $device    
                }
            }
        }
        
      
    }
    
}

$excel.Workbooks.Close()

$wb = $excel.workbooks.open("C:\Users\N.Artykaly\Documents\git\san-zoning\Targets.xlsx")

ForEach ($zoneName in $zones) {
    for ($i=1; $i -le $wb.sheets.count; $i++){
        $sh=$wb.Sheets.Item($i)
    
        if ($sh.Name -eq $zoneName.RightName) {
            $Lines = $sh.UsedRange.Rows.Count
    
            for ($j = 2; $j -le $Lines; $j++) {
                $DeviceName = $sh.Rows.Item($j).columns.Item(1).Text
                $HBA        = $sh.Rows.Item($j).columns.Item(2).Text
                $HBAPort    = $sh.Rows.Item($j).columns.Item(3).Text
                $Fabric     = $sh.Rows.Item($j).columns.Item(4).Text
                $WWN_DN     = $sh.Rows.Item($j).columns.Item(5).Text
                $WWN_DP     = $sh.Rows.Item($j).columns.Item(6).Text
                
                if (-not ([string]::IsNullOrEmpty($Fabric))) {
                    if (! $fabrics.ContainsKey($Fabric)) {
                        $fabrics.Add($Fabric, $Fabric)
                    }

                    $device = [Device]::New("T", $DeviceName, $HBA, $HBAPort, $Fabric, $WWN_DN, $WWN_DP)
    
                    $targets += $device    
                }
            }
        }
        
      
    }
    
}

$excel.Workbooks.Close()

$excel.Quit()

$fabrics.Keys | ForEach-Object {
    $FabricName = $fabrics[$_]
    "Commands for Fabric {0}" -f $FabricName
    "cfgcreate `"MainZoning`""

    $initiators_in_fabric = @()
    $targets_in_fabric = @()

    foreach ($initiator in $initiators) {
        if ($initiator.Fabric -eq $FabricName) {
            $initiators_in_fabric += $initiator
        }
    }

    foreach ($target in $targets) {
        if ($target.Fabric -eq $FabricName) {
            $targets_in_fabric += $target
        }
    }

    foreach ($initiator in $initiators_in_fabric) {
        $initiator.GetCreateAlias()
    }

    foreach ($target in $targets_in_fabric) {
        $target.GetCreateAlias()
    }

    foreach ($zone in $zones) {
        $initiatorName = $zone.LeftName
        $targetName = $zone.RightName

        $initiators_in_zone = @()
        $targets_in_zone = @()

        foreach ($initiator in $initiators_in_fabric) {
            if ($initiator.DeviceName -eq $initiatorName) {
                $initiators_in_zone += $initiator
            }
        }

        foreach ($target in $targets_in_fabric) {
            if ($target.DeviceName -eq $targetName) {
                $targets_in_zone += $target
            }
        }
        
        foreach ($initiator in $initiators_in_zone) {
            foreach ($target in $targets_in_zone) {
                "zonecreate `"{0}_{1}`", `"{0}; {1}`"" -f $initiator.GetAliasName(), $target.GetAliasName()
                "cfgadd `"MainZoning`", `"{0}_{1}`"" -f $initiator.GetAliasName(), $target.GetAliasName()
            }
        }
    }
}


