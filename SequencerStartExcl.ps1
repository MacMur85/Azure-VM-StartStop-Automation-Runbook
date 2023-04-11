Param(
    [parameter(Mandatory=$true)]
    [String] $subscriptionId,
    [parameter(Mandatory=$true)]
    [String] $accountId
)

Import-Module PSDates

$errorActionPreference = "Stop"

# Globals used to collect starting / stopping VMs pools
$global:vmJobListStart = @()
$global:startingVms = @()

# Constants
$STOP_TAG = "automated_stop"
$START_TAG = "automated_start"

$STOP_SEQ_TAG = "sequence_stop"
$START_SEQ_TAG = "sequence_start"

$EXCLUSION_TAG = "automated_excl"

function AddToStartStopLists{
    param(
        [Object]$VirtualMachine,
        [string]$DesiredState
    )

    Write-Output "[$($VirtualMachine.Name)]: Managing VM PowerState..."

    # Retrieve VM with current status
    $eachVM = Get-AzVM -ResourceGroupName $VirtualMachine.ResourceGroupName -Name $VirtualMachine.Name -Status
    $currentStatus = $eachVM.Statuses | Where-Object Code -like "PowerState*" 
    $currentStatus = $currentStatus.Code -replace "PowerState/",""

    # If should be started and isn't, add to starting VM pool
    if($DesiredState -eq "Started" -and $currentStatus -notmatch "running")
    {        
		Write-Output "[$($VirtualMachine.Name)]: Adding VM to STARTING VMs pool"
        $startVmData = Get-AzResource -ResourceType "Microsoft.Compute/VirtualMachines" -Name $VirtualMachine.Name
		$startVmData | Add-Member -NotePropertyName SequenceStart -NotePropertyValue $([int]$($startVmData.Tags.$START_SEQ_TAG)) -Force
        $global:startingVms += $startVmData
    }
	
    # Otherwise, current power state is correct
    else
    {
        Write-Output "[$($VirtualMachine.Name)]: Current state [$currentStatus] is correct."
    }
}

$currentTimeUTC = (Get-Date).ToUniversalTime()

Write-Output "Runbook started..."
Write-Output "Current UTC time: [$($currentTimeUTC.ToString("dddd, yyyy MMM dd HH:mm:ss"))] "

$currentTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($currentTimeUTC, 'GMT Standard Time')

Write-Output "Current GMT Standard Time: [$($currentTime.ToString("dddd, yyyy MMM dd HH:mm:ss"))] "


try
{
    "`nLogging in to Azure..."

    Connect-AzAccount -Identity -AccountId "$accountId" -Subscription "$subscriptionId" | Out-Null

    "Login successful.."
}

catch {
    Write-Error -Message $_.Exception
    throw $_.Exception
}

try {

    $context = Get-AzSubscription -SubscriptionId $subscriptionId
    Set-AzContext $context

	$subName = (Get-AzContext).Subscription.Name
	Write-Output "`nCurrent subscription context: [$($subName)]"

    # Get VM Data
    $allVMs = Get-AzResource -ResourceType "Microsoft.Compute/VirtualMachines" -TagName "$STOP_TAG"

    foreach ($vm in $allVMs) {
        Write-Output "`nProcessing VM: [$($vm.Name)]"

        # Retrieve VMs exclusion status (none, stop, start, both)
        $exclusion_value = $($vm.Tags)."$EXCLUSION_TAG"
        #$no_exclusion = ![string]::IsNullOrWhiteSpace($exclusion_value)

        if ($exclusion_value -ne "both") {

            # Retrieve the VM start and shutdown schedule
            $stop_schedule = $($vm.Tags)."$STOP_TAG"
            $start_schedule = $($vm.Tags)."$START_TAG"

            $stop_sequence = $($vm.Tags)."$STOP_SEQ_TAG"
            $start_sequence = $($vm.Tags)."$START_SEQ_TAG"

            # Start schedules are optional on VMs
            $has_start_schedule = ![string]::IsNullOrWhiteSpace($start_schedule)
            
            # Sequencers are optional
            $has_stop_sequencer = ![string]::IsNullOrWhiteSpace($stop_sequence)
            $has_start_sequencer = ![string]::IsNullOrWhiteSpace($start_sequence)

            if([string]::IsNullOrWhiteSpace($stop_schedule))
            {
                Write-Output "[$($vm.Name)]: Skipping due to invalid/empty value in $STOP_TAG"
                continue
            }

            if((Test-CrontabSchedule -crontab "$stop_schedule").Valid -eq $false)
            {
                Write-Output "[$($vm.Name)]: Skipping this VM due to invalid cron schedule: $stop_schedule"
                continue
            }

            if($has_start_schedule -eq $true) {

                if((Test-CrontabSchedule -crontab "$start_schedule").Valid -eq $false) {
                    Write-Output "[$($vm.Name)]: Skipping this VM due to invalid cron schedule: $start_schedule"
                }
                
                if($has_start_sequencer -eq $true) {
                    # verify start sequence
                    $parsedStart = ""
                    if(![Int]::TryParse($start_sequence,[ref]$parsedStart)) 
                    {
                        Write-Output "[$($vm.Name)]: Skipping due to non-numeric or NULL value in $START_SEQ_TAG"
                        continue
                    }
                }

                Write-Output "[$($vm.Name)]: Start schedule is: $start_schedule"
            }
            Write-Output "[$($vm.Name)]: Stop schedule is: $stop_schedule"

            # See what the target stop time is for the VM on current schedule
            $nextStopTime = Get-CronNextOccurrence "$stop_schedule"  -StartTime $currentTime.Date # -EndTime $currentTime.Date.AddDays(1)
            Write-Output "[$($vm.Name)]: Target stop time is: $nextStopTime"

            # If VM has a start time and a stop time
            if ($has_start_schedule -eq $true) {
                
                # See what the target start time is for the VM on current schedule
                $nextStartTime = Get-CronNextOccurrence "$start_schedule"  -StartTime $currentTime.Date # -EndTime $currentTime.Date.AddDays(1)
                Write-Output "[$($vm.Name)]: Target start time is: $nextStartTime"

                # If we are between the start and stop times the VM should be running
                if (($currentTime -ge $nextStartTime) -And ($currentTime -le $nextStopTime)) {
                    if($has_start_sequencer -eq $true) {

                        # If VM has a start exclusion
                        if ($exclusion_value -eq "start") {
                            Write-Output "[$($vm.Name)]: Start Exclusion detected. Skipping VM due to '$($EXCLUSION_TAG) : $($exclusion_value)' tag..."
                            continue
                        }
                        else {                      
                            AddToStartStopLists -VirtualMachine $vm -DesiredState "Started"
                        }
                    }
                    else {
                        Write-Output "[$($vm.Name)]: No $START_SEQ_TAG tag or value, another Runbook will handle this VM. Skipping..."
                    }
                } 
                
                else {
                    Write-Output "[$($vm.Name)]: Stop action detected, another Runbook will handle this VM. Skipping..."
                }
            } 
            
            else {

                # If VM just has a stop time defined
                if (($currentTime -ge $nextStopTime)) {
                    Write-Output "[$($vm.Name)]: Stop action detected, another Runbook will handle this VM. Skipping..."
                }
            }
        }
        elseif ($exclusion_value -eq "both") {
            Write-Output "[$($vm.Name)]: Start/Stop Exclusion detected. Skipping VM due to '$($EXCLUSION_TAG) : $($exclusion_value)' tag..."
            continue
        }  
    }
    
	
	# Sorting START list of VMs by sequencer 
	$global:startingVms = $global:startingVms | Sort-Object -Property SequenceStart

	# Write-Output "Verifying STOP and START pools..."
    Write-Output "`nVerifying START VMs pool..."

    if (($global:startingVms | Measure-Object).Count -gt 0) {

	    Write-Output "Starting VMs in below sequence:"
	    Write-Output $($global:startingVms.Name)

        foreach ($vm in $global:startingVms) {

            Write-Output "[$($vm.Name)] Turning VM on..."
            $global:vmJobListStart += $vm | Start-AzVM 
    	}
	}

    else {
        Write-Output "`nNo VMs to start"
    }

}
catch {
    Write-Error "$($_.Exception.Message)"
    throw "Error executing stop/start script"
}
finally {
	
    Write-Output "`nRunbook finished (Duration: $(("{0:hh\:mm\:ss}" -f ((Get-Date).ToUniversalTime() - $currentTime))))"
}
