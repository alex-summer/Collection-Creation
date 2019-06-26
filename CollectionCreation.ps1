#Script that automates the creation of collections for our laptop carts. Also known as COWS.


#Changing to SCCM site
Import-Module -Name "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
cd #SCCM Site

#Defining variables
$SourceData = Import-Csv -Path "$env:userprofile\Desktop\COWINFO.csv" -Delimiter "," | Where-Object{$_.'SCCM Collection' -ne "yes"}
$Cows= @()
$i = 0 #Using "$i" as a counter to keep an even distribution of schedules across the collections.
$Model = @(<#Array of models. Used this to put the correct model in the collection comments #>)
$Location = @(<#Array of locations. Used this to put collection in the correct container by matching what was put into the csv folder#>)
$Schedule = @(
    New-CMSchedule -Start '06/01/2019 9:00 PM' -DayofWeek Monday -RecurCount 1
    New-CMSchedule -Start '06/01/2019 9:00 PM' -DayofWeek Tuesday -RecurCount 1
    New-CMSchedule -Start '06/01/2019 9:00 PM' -DayofWeek Wednesday -RecurCount 1
    New-CMSchedule -Start '06/01/2019 9:00 PM' -DayofWeek Thursday -RecurCount 1
    New-CMSchedule -Start '06/01/2019 9:00 PM' -DayofWeek Friday -RecurCount 1
    New-CMSchedule -Start '06/01/2019 9:00 PM' -DayofWeek Saturday -RecurCount 1
    New-CMSchedule -Start '06/01/2019 9:00 PM' -DayofWeek Sunday -RecurCount 1
    )

#Put each cow into an array of objects
foreach ($Source in $SourceData)
{
    
    $Cow = New-Object -TypeName PSObject -Property @{
        Name = "COW$($Source.'COW#')"
        Comment = "Count: $($Source.Count) // $($Model -match $Source.Model)"
        Destination = ($Location -match $Source.Destination)[0]
        SCCMCollection = $Source.'SCCM Collection'
        Query = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,`
            SMS_R_SYSTEM.SMSUniqueIdentifier,`
            SMS_R_SYSTEM.ResourceDomainORWorkgroup,`
            SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.Name like `"COW$($Source.'COW#')-%`""
        Schedule = $Schedule[$i]
    }

    $Cows += $Cow

    if($i -eq 7){$i = 0}
    else{$i++}
}

#All the heavy lifting here (Collection creation, adds membership rules and move the collection to the right place)
foreach ($Cow in $Cows)
{    
    $collection = New-CMDeviceCollection -Name "$($Cow.Name)" -Comment "$($Cow.Comment)" -LimitingCollectionName "All Systems" -RefreshType Periodic -RefreshSchedule $Cow.Schedule -whatif
    Add-CMDeviceCollectionQueryMembershipRule -Collection $Collection -QueryExpression $Cow.Query -RuleName "Name" -whatif 
    Move-CMObject -InputObject $Collection -FolderPath "Path to destination\$($Cow.Destination)" -whatif
}
