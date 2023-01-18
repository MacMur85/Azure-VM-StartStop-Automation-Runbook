# Azure-VM-StartStop-Automation-Runbook
Azure PowerShell scripts for Automation Account Runbooks for automated VM Start/Stop with sequencers and exclusions

Those three script start and stop VMs across a subscription based on cron schedules specified tags on the VMs.
For example:

`automation_start : * 8 * * 1-5`
and
`automation_stop : * 18 * * 1-5`

Current Functionality:

- Start and stop VMs based on tags
- Start and stop VMs based on a sequence tag on the resource. For example: `sequence_start : 2`, `sequence_stop : 2`
- Exclusions from automation based on an exclusion tag and its value. For example:
  - exclude from automation: `automated_excl : both` 
  - exclude from automated startup: `automated_excl : start`  
  - exclude from automated shut down: `automated_excl : stop`
