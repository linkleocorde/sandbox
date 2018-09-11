$starttime = get-date
$reportPath = 'E:\Temp\v1EnterpriseEpics.csv'

$dt = New-Object System.Data.DataTable

$col = New-Object System.Data.DataColumn("Name", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Number", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("PlannedStart", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("PlannedEnd", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("oid", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Scope.Name", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Scope.Number", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Status.Name", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Category.Name", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Team.Name", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Team.Number", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("ActiveWorkitems", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("ClosedWorkitems", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Epic.Name", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Epic.Number", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Epic.Status.Name", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Epic.Scope.Name", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Epic.Category.Name", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Super.Name", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Super.Number", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Super.Status.Name", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Super.Scope.Name", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("Super.Category.Name", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("APIQueryDate", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("CreateDate", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("ClosedDate", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("AssetState", [string])
$dt.Columns.Add($col)
$col = New-Object System.Data.DataColumn("IsDeleted", [string])
$dt.Columns.Add($col)


#Generate VersionOne Epic Reports
$reqHeaders = @{}
$reqHeaders.Add("ServerHost", "www10.v1host.com")
$reqHeaders.Add("Authorization", "Bearer XXXXXXXXXXXXXXXXXXXXXXXXXX")
$PostUri = 'https://www10.v1host.com/NIKE01a/query.v1'

$pepicQuery = @"
from: Epic
select:
    - Name
    - Number
    - Team
    - PlannedStart
    - PlannedEnd
    - Team.Name
    - Scope.Name
    - Scope
    - Status.Name
    - Category.Name
    - CreateDate
    - ClosedDate
    - IsDeleted
    - AssetState
    - Super.Name
    - Super.Number
    - Super.Scope.Name
    - Super.Status.Name
    - Super.Category.Name
    - Subs
    - SubsAndDown:PrimaryWorkitem[AssetState!='Dead'].@Count
    - SubsAndDown:PrimaryWorkitem[AssetState='Closed'].@Count
where:
    Category.Name: Enterprise Epic
"@


$pepicResults = Invoke-RestMethod -Uri $PostUri -Body $pepicQuery -Headers $reqHeaders -Method POST -ContentType 'application/xml'

foreach($pepic in $pepicResults[0]){
    if($pepic.Subs.Count -gt 0){
        $epiccount = 0
        foreach($sub in $pepic.Subs){
            if($sub._oid -match 'Epic'){
                $sub._oid
                $epiccount++
                $epicQuery = @"
                from: Epic
                select:
                    - Name
                    - Number
                    - Scope.Name
                    - Status.Name
                    - Category.Name
                where:
                    ID: $($sub._oid)

"@
                $epicResults = Invoke-RestMethod -Uri $PostUri -Method POST -Body $epicQuery -Headers $reqHeaders -ContentType 'application/xml'

                $newrow = $dt.NewRow()
                $newrow.'Epic.Number' = $epicResults[0].Number
                $newrow.'Epic.Name' = $epicResults[0].Name
                $newrow.'Epic.Category.Name' = $epicResults[0].'Category.Name'
                $newrow.'Epic.Scope.Name' = $epicResults[0].'Scope.Name'
                $newrow.'Epic.Status.Name' = $epicResults[0].'Status.Name'
                $newrow.Name = $pepic.Name
                $newrow.Number = $pepic.Number
                $newrow.oid = $pepic._oid
                $newrow.'Status.Name' = $pepic.'Status.Name'
                $newrow.'Scope.Name' = $pepic.'Scope.Name'
                $newrow.'Scope.Number' = $pepic.Scope._oid
                $newrow.'Category.Name' = $pepic.'Category.Name'
                $newrow.ActiveWorkitems = $pepic."SubsAndDown:PrimaryWorkitem[AssetState!='Dead'].@Count"
                $newrow.ClosedWorkitems = $pepic."SubsAndDown:PrimaryWorkitem[AssetState='Closed'].@Count"
                $newrow.'Super.Name' = $pepic.'Super.Name'
                $newrow.'Super.Number' = $pepic.'Super.Number'
                $newrow.'Super.Scope.Name' = $pepic.'Super.Scope.Name'
                $newrow.'Super.Status.Name' = $pepic.'Super.Status.Name'
                $newrow.'Super.Category.Name' = $pepic.'Super.Category.Name'
                $newrow.APIQueryDate = $starttime
                $newrow.CreateDate = Get-Date $pepic.CreateDate -Format g
                if($pepic.ClosedDate -ne $null){
                    $newrow.ClosedDate = Get-Date $pepic.ClosedDate -Format g
                }
                if($pepic.PlannedStart -ne $null){
                    $newrow.PlannedStart = Get-Date $pepic.PlannedStart -Format g
                }
                if($pepic.PlannedEnd -ne $null){
                    $newrow.PlannedEnd = Get-Date $pepic.PlannedEnd -Format g
                }
                $newrow.AssetState = $pepic.AssetState
                $newrow.IsDeleted = $pepic.IsDeleted
                $newrow.'Team.Name' = $pepic.'Team.Name'
                $newrow.'Team.Number' = $pepic.Team._oid



                $dt.rows.Add($newrow)
            }
        }
        if($epiccount -lt 1){
            $newrow = $dt.NewRow()
            $newrow.Name = $pepic.Name
            $newrow.Number = $pepic.Number
            $newrow.oid = $pepic._oid
            $newrow.'Status.Name' = $pepic.'Status.Name'
            $newrow.'Scope.Name' = $pepic.'Scope.Name'
            $newrow.'Scope.Number' =$pepic.Scope._oid
            $newrow.'Category.Name' = $pepic.'Category.Name'
            $newrow.ActiveWorkitems = $pepic."SubsAndDown:PrimaryWorkitem[AssetState!='Dead'].@Count"
            $newrow.ClosedWorkitems = $pepic."SubsAndDown:PrimaryWorkitem[AssetState='Closed'].@Count"
            $newrow.'Super.Name' = $pepic.'Super.Name'
            $newrow.'Super.Number' = $pepic.'Super.Number'
            $newrow.'Super.Scope.Name' = $pepic.'Super.Scope.Name'
            $newrow.'Super.Status.Name' = $pepic.'Super.Status.Name'
            $newrow.'Super.Category.Name' = $pepic.'Super.Category.Name'
            $newrow.CreateDate = Get-Date $pepic.CreateDate -Format g
            if($pepic.ClosedDate -ne $null){
                $newrow.ClosedDate = Get-Date $pepic.ClosedDate -Format g
            }
            if($pepic.PlannedStart -ne $null){
                $newrow.PlannedStart = Get-Date $pepic.PlannedStart -Format g
            }
            if($pepic.PlannedEnd -ne $null){
                $newrow.PlannedEnd = Get-Date $pepic.PlannedEnd -Format g
            }
            $newrow.AssetState = $pepic.AssetState
            $newrow.APIQueryDate = $starttime
            $newrow.IsDeleted = $pepic.IsDeleted
            $newrow.'Team.Name' = $pepic.'Team.Name'
            $newrow.'Team.Number' = $pepic.Team._oid


            $dt.rows.Add($newrow)
        }
    }
    else{
        $newrow = $dt.NewRow()
        $newrow.Name = $pepic.Name
        $newrow.Number = $pepic.Number
        $newrow.oid = $pepic._oid
        $newrow.'Status.Name' = $pepic.'Status.Name'
        $newrow.'Scope.Name' = $pepic.'Scope.Name'
        $newrow.'Scope.Number' =$pepic.Scope._oid
        $newrow.'Category.Name' = $pepic.'Category.Name'
        $newrow.ActiveWorkitems = $pepic."SubsAndDown:PrimaryWorkitem[AssetState!='Dead'].@Count"
        $newrow.ClosedWorkitems = $pepic."SubsAndDown:PrimaryWorkitem[AssetState='Closed'].@Count"
        $newrow.'Super.Name' = $pepic.'Super.Name'
        $newrow.'Super.Number' = $pepic.'Super.Number'
        $newrow.'Super.Scope.Name' = $pepic.'Super.Scope.Name'
        $newrow.'Super.Status.Name' = $pepic.'Super.Status.Name'
        $newrow.'Super.Category.Name' = $pepic.'Super.Category.Name'
        $newrow.APIQueryDate = $starttime
        $newrow.CreateDate = Get-Date $pepic.CreateDate -Format g
        if($pepic.ClosedDate -ne $null){
            $newrow.ClosedDate = Get-Date $pepic.ClosedDate -Format g
        }
        if($pepic.PlannedStart -ne $null){
            $newrow.PlannedStart = Get-Date $pepic.PlannedStart -Format g
        }
        if($pepic.PlannedEnd -ne $null){
            $newrow.PlannedEnd = Get-Date $pepic.PlannedEnd -Format g
        }
        $newrow.AssetState = $pepic.AssetState
        $newrow.IsDeleted = $pepic.IsDeleted
        $newrow.'Team.Name' = $pepic.'Team.Name'
        $newrow.'Team.Number' = $pepic.Team._oid

        $dt.rows.Add($newrow)
    }

}

"Exporting $reportPath"
$dt | Export-Csv -Path $reportPath -NoTypeInformation

$endtime = get-date
$runtime = $endtime - $starttime
"Runtime: $runtime"
