#Param(
#    [parameter(Mandatory=$true)]
#    [String] $subscriptionId,
#    [parameter(Mandatory=$true)]
#    [String] $accountId
#)

Import-Module PSDates

$errorActionPreference = "Stop"

$subscriptionId = "9e15b045-0efb-4289-a18d-1a11052f1068"
$accountId = "ca253b5f-34a8-4917-88e6-72f1617bebe1"

# Globals used to collect starting / stopping VMs pools
$global:vmJobListStop = @()
$global:stoppingVms = @()

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

    # If should be stopped and isn't, add to stopping VM pool
    if($DesiredState -eq "StoppedDeallocated" -and $currentStatus -ne "deallocated")
    {
        Write-Output "[$($VirtualMachine.Name)]: Adding VM to STOPPING VMs pool"
		$stopVmData = Get-AzResource -ResourceType "Microsoft.Compute/VirtualMachines" -Name $VirtualMachine.Name
		$stopVmData | Add-Member -NotePropertyName SequenceStop -NotePropertyValue $([int]$($stopVmData.Tags.$STOP_SEQ_TAG)) -Force
        $global:stoppingVms += $stopVmData
    }
	
    # Otherwise, current power state is correct
    else
    {
        Write-Output "[$($VirtualMachine.Name)]: Current state [$currentStatus] is correct."
    }
}

$currentTime = (Get-Date).ToUniversalTime()

Write-Output "Runbook started..."
Write-Output "Current UTC time: [$($currentTime.ToString("dddd, yyyy MMM dd HH:mm:ss"))] "

try
{
    "`nLogging in to Azure..."

    Connect-AzAccount -Identity -AccountId "$accountId" -Subscription "$subscriptionId" | Out-Null

    "Login successful..."
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
            

            # Start schedules are optional on VMs
            $has_start_schedule = ![string]::IsNullOrWhiteSpace($start_schedule)
            
            # Sequencers are optional
            $has_stop_sequencer = ![string]::IsNullOrWhiteSpace($stop_sequence)
            #$has_start_sequencer = ![string]::IsNullOrWhiteSpace($start_sequence)

            if([string]::IsNullOrWhiteSpace($stop_schedule)) {
                Write-Output "[$($vm.Name)]: Skipping due to invalid/empty value in $STOP_TAG"
                continue
            }

            if((Test-CrontabSchedule -crontab "$stop_schedule").Valid -eq $false) {
                Write-Output "[$($vm.Name)]: Skipping this VM due to invalid cron schedule: $stop_schedule"
                continue
            }

            if($has_stop_sequencer -eq $true) {
                # Verify stop sequencer
                $parsedStop = ""
                if(![Int]::TryParse($stop_sequence,[ref]$parsedStop)) {
                    Write-Output "[$($vm.Name)]: Skipping due to non-numeric or NULL value in $STOP_SEQ_TAG"
                    continue
                }
            }

            Write-Output "[$($vm.Name)]: Stop schedule is: $stop_schedule"

            # See what the target stop time is for the VM on current schedule
            $nextStopTime = Get-CronNextOccurrence "$stop_schedule"  -StartTime $currentTime.Date -EndTime $currentTime.Date.AddDays(1)
            Write-Output "[$($vm.Name)]: Target stop time is: $nextStopTime"

            # If VM has a start time and a stop time
            if ($has_start_schedule -eq $true) {
                
                # See what the target start time is for the VM on current schedule
                $nextStartTime = Get-CronNextOccurrence "$start_schedule"  -StartTime $currentTime.Date -EndTime $currentTime.Date.AddDays(1)
                Write-Output "[$($vm.Name)]: Target start time is: $nextStartTime"

                # If we are between the start and stop times the VM should be running
                if (($currentTime -ge $nextStartTime) -And ($currentTime -le $nextStopTime)) {
                    Write-Output "[$($vm.Name)]: Sequenced Start action detected, another Runbook will handle this VM. Skipping..."
                } 
                
                else {
                    if($has_stop_sequencer -eq $true) {

                        # If VM has a stop exclusion
                        if ($exclusion_value -eq "stop") {
                            Write-Output "[$($vm.Name)]: Stop Exclusion detected. Skipping VM due to '$($EXCLUSION_TAG) : $($exclusion_value)' tag..."
                            continue
                        }
                        else {
                            AddToStartStopLists -VirtualMachine $vm -DesiredState "StoppedDeallocated"
                        }
                    }           
                    else {
                        Write-Output "[$($vm.Name)]: No $STOP_SEQ_TAG tag or value, another Runbook will handle this VM. Skipping..."
                    }
                    
                }
            }
            else { 

                # If VM just has a stop time defined
                if (($currentTime -ge $nextStopTime)) {

                    # If VM has a stop exclusion
                    if($has_stop_sequencer -eq $true) {
                        if ($exclusion_value -eq "stop") {
                            Write-Output "[$($vm.Name)]: Stop Exclusion detected. Skipping VM due to '$($EXCLUSION_TAG) : $($exclusion_value)' tag..."
                            continue
                        }
                        else {    
                            AddToStartStopLists -VirtualMachine $vm -DesiredState "StoppedDeallocated"
                        }
                    }
                    else {
                        Write-Output "[$($vm.Name)]: No $STOP_SEQ_TAG tag or value, another Runbook will handle this VM. Skipping..."
                    }
                }
            }

        }
        elseif ($exclusion_value -eq "both") {
            Write-Output "[$($vm.Name)]: Start/Stop Exclusion detected. Skipping VM due to '$($EXCLUSION_TAG) : $($exclusion_value)' tag..."
            continue
        } 
    }
    
    # Sorting STOP list of VMs by sequencer 
	$global:stoppingVms = $global:stoppingVms | Sort-Object -Property SequenceStop

	# Write-Output "Verifying STOP and START pools..."
    Write-Output "`nVerifying STOP VMs pool..."

    if (($global:stoppingVms | Measure-Object).Count -gt 0) {
    
        Write-Output "Stopping VMs in below sequence:"
        Write-Output $($global:stoppingVms.Name)

        foreach ($vm in $global:stoppingVms) {
            
            Write-Output "[$($vm.Name)]: Turning VM off..."
            $global:vmJobListStop += $vm | Stop-AzVM -Force
            }
    }

    else {
        Write-Output "`nNo VMs to stop"
    }

}
catch {
    Write-Error "$($_.Exception.Message)"
    throw "Error executing stop/start script"
}
finally {
	
    Write-Output "`nRunbook finished (Duration: $(("{0:hh\:mm\:ss}" -f ((Get-Date).ToUniversalTime() - $currentTime))))"
}