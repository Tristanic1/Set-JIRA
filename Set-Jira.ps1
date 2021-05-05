#requires -Modules JIRAPS
#requires -Modules CredentialManager

function Set-Ticket
{
    <#
        .SYNOPSIS
        Function takes unassigned ticket, updates Category, Service, Assignee fields.
		Information about changes is saved to the log file.

        .DESCRIPTION
        Function takes unassigned ticket, updates Category, Service, Assignee fields.
		Information about changes is saved to the log file.

        .PARAMETER InputObject
        Unassigned ticket (usually from the Pipeline)

        .PARAMETER Values
        Defines Category, Service and Assignee changes.

        .EXAMPLE
        Assign-Ticket -InputObject $JIRAIssue -Values $Phones
        
        .INPUTS
		[Object]
		
		.OUTPUTS
		Ticket modified in JIRA.
        Saves successfull assigment to the log file.
    #>

	param
	(
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Ticket to assign')]
		[Object]$InputObject,
		[Parameter(Mandatory=$true,HelpMessage='Hashtable name with Category/Service/Assignee values')]
		[hashtable]$Values
	)
	process
	{
		Set-JiraIssue @Values -Issue $InputObject -Credential $Cred
		
		# This will create line like this in the console and in logfile: 
		# [2020-05-20 11:20:47] Assigned Issue SD-101122 Newcomer starting 1st May --> Smith 
		$Result = '[{0}] Assigned Issue {1} {2} --> {3}' -f ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss"), $InputObject.Key, $InputObject.Summary, $Values['Assignee']
		$Result
		Add-Content -Path $LogFile -Value $Result

	}
}

function Set-Facility
{
    param
    (
        [Object]
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Tickets for Facility to be assigned")]
        $InputObject
    )
    process
    {
		
		switch -Regex ($_.Summary) {
			# Furniture requests
			'drawer|cabinet|chair|furniture|desk\b|whiteboard|lamp|coat|sideboard|cupboard' {
				Set-Ticket -InputObject $InputObject -Values $Furniture
				break
			}
			# Table (furniture) requests - tricky, because can be SQL Table, Signature Table in document, Pivot Table, Table of Contents etc.
			'table' {
				if (($_.Summary -Notmatch 'SQL|Database|contents|pivot|filter|payrol|query|signature|phone') -and ($_.Reporter -notmatch 'automatedSystem')) {
					Set-Ticket -InputObject $InputObject -Values $Furniture
					break
				}				
			} 
			# Newcomers requests
			'FACILITY$' {
				Set-Ticket -InputObject $InputObject -Values $Newcomer
				break
			}
			# Office move and accommodation requests
			'office move|accommodation|moving box' {
				Set-Ticket -InputObject $InputObject -Values $OfficeMove
				break
			}
			# HVAC requests: Air conditioning, radiator, heating, cooling
			'condition|radiator|heat|cool' {
				Set-Ticket -InputObject $InputObject -Values $HVAC
				break
			}
			# Building repair requests
			'window\b|door\b|blind|hinge|light|toilet|hot water|water tap|leak' {
				Set-Ticket -InputObject $InputObject -Values $BuildingMaintenance
				break
			}
			# Issues with electricity in the office
			'power' {
				if ($_.Summary -notmatch 'cord|cable|supply|adapter') {
					Set-Ticket -InputObject $InputObject -Values $PowerIssues
					break
				}				
			}
			# Cleaning requests
			'cleaning|d[ei]sinfect|soap|sanitizer|vacuum' {
				Set-Ticket -InputObject $InputObject -Values $Cleaning
				break
			}
			# Printer alerts for Facility
			'^PRINTER-ALERT:[WC]|^PRINTER-ALERT:TO|drucker' {
				Set-Ticket -InputObject $InputObject -Values $PrinterFacility
				break
			}
			# Printer alert for Facility: toner threshold
			'^PRINTER-ALERT:TH' {
				if ($_.Summary -Match 'toner') { Set-Ticket -InputObject $InputObject -Values $PrinterFacility }
				break
			}

			# Reception services
			'photo|picture|door label' {
				Set-Ticket -InputObject $InputObject -Values $ReceptionServices
				break
			}
			# Access Badge requests
			'access to datacent|room access|access badge|badge access' {
				Set-Ticket -InputObject $InputObject -Values $FacilityBadges
				break
			}
			# Key and locker requests
			'key' {
				if ($_.Summary -Match 'office|drawer|cupboard|locker|door\s') { Set-Ticket -InputObject $InputObject -Values $FacilityKeys }
				break
			}
			# Waste disposal request
			'container|trash|waste|disposal' {
				Set-Ticket -InputObject $InputObject -Values $WasteDisposal
				break
			}
			# Expedition and delivery requests
			'warehouse|palette|\bboxes\b' {
				Set-Ticket -InputObject $InputObject -Values $Expedition
				break
			}
			# VIP driving requests
			'drive VIP' {
				Set-Ticket -InputObject $InputObject -Values $VIPdriving
				break
			}
		}
        if ($_.Reporter -Match 'security_guards') { Set-Ticket -InputObject $InputObject -Values $BuildingMaintenance }
	}
}


# Connect to JIRA Server 
Set-JiraConfigServer -Server 'https://jira.companyname.com'

# Takes stored credentials from Credential Manager on local computer (Windows Credential -> Generic Credential -> JIRA)
# JIRA requires authentication for each action, at minimum you will need to use "Get-Credential"
$cred = Get-StoredCredential -Target JIRA

#---------------------------------
# Dispatchers and Default Assignees

# Divisional dispatchers
$FacilityDispatcher = 'Banner'
$ITSupplyDispatcher = 'Romanoff'

# Service Desk assignees
$ITRequestSystem = 'Rogers'
$Hardware = 'Parker'
$WebExSupport = 'Stark'
$Phones = 'Stark'

#---------------------------------

# Creating a log file
$Date = Get-Date -Format FileDate
$Filename = $date+'.log'
$LogFile = Join-Path $PSScriptRoot ($Filename)

if (!(Test-Path -Path $LogFile)) {
	New-Item -Path $PSScriptRoot -Name $Filename > $null
}
#---------------------------------

# Defined the variables for Category, Service and Assignees
#-------- Facility Stuff --------
$PrinterFacility = @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Copier/Printer Issues'
				'child' = @{'value'='Copier/Printer Supplies'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Copier/Printer Issues'}
		}
	 }
}

$Newcomer = @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Accomodation Issues'
				'child' = @{'value'='Newcomers'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Accommodation Issues'}
		}
	 }
}

$BuildingMaintenance = @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Building Maintenance'
				'child' = @{'value'='Repair Requirement'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Building Maintenance'}
		}
	 }
}

$Furniture = @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Accomodation Issues'
				'child' = @{'value'='Furniture Request'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Accommodation Issues'}
		}
	 }
}

$OfficeMove = @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Accomodation Issues'
				'child' = @{'value'='Office Move Request'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Accommodation Issues'}
		}
	 }
}

$HVAC = @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Building Maintenance'
				'child' = @{'value'='HVAC'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Building Maintenance'}
		}
	 }
}

$PowerIssues = @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Building Maintenance'
				'child' = @{'value'='Power Supply Issues'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Building Maintenance'}
		}
	 }
}

$Cleaning = @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Cleaning Issues'
				'child' = @{'value'='Requests'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Cleaning Issues'}
		}
	 }
}

$ReceptionServices = @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Reception Services'
				'child' = @{'value'='Phone Book & Door Labels'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Reception Services'}
		}
	 }
}

$FacilityBadges = @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Security, Health & Safety Issues'
				'child' = @{'value'='ID Badges'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Security, Health & Safety Issues'}
		}
	 }
}

$FacilityKeys = @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Security, Health & Safety Issues'
				'child' = @{'value'='Keys'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Security, Health & Safety Issues'}
		}
	 }
}

$WasteDisposal= @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Transport & Logistics'
				'child' = @{'value'='Waste Disposal'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Transport & Logistics'}
		}
	 }
}

$Expedition= @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Transport & Logistics'
				'child' = @{'value'='Expedition and Delivery'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Transport & Logistics'}
		}
	 }
}

$VIPdriving= @{
	Assignee = "$FacilityDispatcher"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Facility - Transport & Logistics'
				'child' = @{'value'='VIP Driving Requests'}
		}
		customfield_10201 = @{
				'value' = 'Facility'
				'child' = @{'value'='Transport & Logistics'}
		}
	 }
}

#-------- IT Stuff --------
$WebEx = @{
	Assignee = "$WebExSupport"
	Fields = @{
		customfield_10200 = @{
				'value' = 'IT Admin'
				'child' = @{'value'='Account Creation/Modification'}
		}
		customfield_10201 = @{
				'value' = 'IT'
				'child' = @{'value'='COTS & Bespoke Application'}
		}
	 }
}

$Software = @{
	Assignee = "$ITRequestSystem"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Desktop Software'
				'child' = @{'value'='Other'}
		}
		customfield_10201 = @{
				'value' = 'IT'
				'child' = @{'value'='ITRequestSystem'}
		}
	 }
}

$ServiceAccess = @{
	Assignee = "$ITRequestSystem"
	Fields = @{
		customfield_10200 = @{
				'value' = 'IT Admin'
				'child' = @{'value'='Account Creation/Modification'}
		}
		customfield_10201 = @{
				'value' = 'IT'
				'child' = @{'value'='ITRequestSystem'}
		}
	 }
}

$VPN = @{
	Assignee = "$ITRequestSystem"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Mobile Computing Service'
				'child' = @{'value'='VPN'}
		}
		customfield_10201 = @{
				'value' = 'IT'
				'child' = @{'value'='Mobile Computing'}
		}
	 }
}

$Headset = @{
	Assignee = "$Hardware"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Desktop Hardware'
		}
		customfield_10201 = @{
				'value' = 'IT'
				'child' = @{'value'='Desktop Computing'}
		}
	 }
}

$Monitor = @{
	Assignee = "$Hardware"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Desktop Hardware'
				'child' = @{'value'='Screen'}
		}
		customfield_10201 = @{
				'value' = 'IT'
				'child' = @{'value'='Desktop Computing'}
		}
	 }
}

$Peripheral = @{
	Assignee = "$Hardware"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Desktop Hardware'
				'child' = @{'value'='Keyboard/Mouse'}
		}
		customfield_10201 = @{
				'value' = 'IT'
				'child' = @{'value'='Desktop Computing'}
		}
	 }
}

$laptop = @{
	Assignee = "$Hardware"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Mobile Computing Service'
				'child' = @{'value'='Notebook'}
		}
		customfield_10201 = @{
				'value' = 'IT'
				'child' = @{'value'='Mobile Computing'}
		}
	 }
}

$PrinterIT = @{
	Assignee = "$Hardware"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Printer'
				'child' = @{'value'='Corridor Printer'}
		}
		customfield_10201 = @{
				'value' = 'IT'
				'child' = @{'value'='Print Service'}
		}
	 }
}

$Deskphone = @{
	Assignee = "$Phones"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Telecommunication'
				'child' = @{'value'='Office phone'}
		}
		customfield_10201 = @{
				'value' = 'IT'
				'child' = @{'value'='Telecommunication'}
		}
	 }
}

$iPhone = @{
	Assignee = "$Phones"
	Fields = @{
		customfield_10200 = @{
				'value' = 'Telecommunication'
				'child' = @{'value'='Mobile phone'}
		}
		customfield_10201 = @{
				'value' = 'IT'
				'child' = @{'value'='Telecommunication'}
		}
	 }
}

$JIRA = @{
	Assignee = "$ITSupplyDispatcher"	
	Fields = @{
		customfield_10200 = @{
				'value' = 'Other'
				'child' = @{'value'='Other'}
		}
		customfield_10201 = @{
				'value' = 'IT'
				'child' = @{'value'='JIRA'}
		}
	 }
}
#---------------------------------

# Filter in JIRA: project = SD AND status != Resolved AND assignee in (EMPTY)
$UnassignedIssues = Get-JiraIssue -Filter '12345' -Credential $cred

# Assign the matched Facility tickets
$UnassignedIssues | Set-Facility

# Removing the Facility tickets  that were already assigned in the last operation
$UnassignedIssues = Get-JiraIssue -Filter '12345' -Credential $cred

# Requests for laptop/notebook or any related issues with them 
$UnassignedIssues | Where-Object {$_.Summary -Match 'need laptop|laptop request|laptop for|need notebook|notebook request|notebook for|HARDWARE$'} | 
	Set-Ticket -Values $Laptop		
	
# WebEx/zoom account requests and support
$UnassignedIssues | Where-Object {$_.Summary -Match 'webex|zoom'} |
	Set-Ticket -Values $WebEx

# VPN, token requests/issues
$UnassignedIssues | Where-Object {$_.Summary -Match 'VPN|token'} |
	Set-Ticket -Values $VPN

# "ITRequestSystem: Service Access Request", "ITRequestSystem: Account Extension Request", "ITRequestSystem: Administrative Rights Request"
$UnassignedIssues | Where-Object {$_.Summary -CMatch 'SERVICE ACCESS$|ACCOUNT$|^ITRequestSystem: Account E|^ITRequestSystem: Ad'} | 
	Set-Ticket -Values $ServiceAccess
		
# ITRequestSystem: Software Request	
$UnassignedIssues | Where-Object {$_.Summary -Match '^ITRequestSystem: So'} | 
	Set-Ticket -Values $Software

# Printer alerts for IT: Maintenance, Service Call, Imaging Unit
$UnassignedIssues | Where-Object {$_.Summary -Match '^PRINTER-ALERT:[MNSI]'} | 
	Set-Ticket -Values $PrinterIT

# Printer alerts for IT: imaging unit and drum cartridge threshold
$UnassignedIssues | Where-Object {($_.Summary -Match '^PRINTER-ALERT:TH') -and ($_.Summary -Match 'drum|unit')} | 
	Set-Ticket -Values $PrinterIT

# Headset hardware requests
$UnassignedIssues | Where-Object {$_.Summary -Match 'headset|headphone'} | 
	Set-Ticket -Values $Headset

# Keyboard/mouse requests
$UnassignedIssues | Where-Object {$_.Summary -Match 'keyboard|mouse|webcam'} | 
	Set-Ticket -Values $Peripheral

# Requests for screen/monitor or any related issues with them 
$UnassignedIssues | Where-Object {$_.Summary -Match 'monitor\s|broken screen|second screen|2nd screen'} | 
	Set-Ticket -Values $Monitor

# Requests for deskphones or any related issues with them 
$UnassignedIssues | Where-Object {$_.Summary -Match 'TELEPHONE$|deskphone|desk phone'} | 
	Set-Ticket -Values $Deskphone

# Requests for iphones or any related issues with them 
$UnassignedIssues | Where-Object {$_.Summary -Match 'iphone|mobile|loan phone'} | 
	Set-Ticket -Values $iPhone

# Tickets mentioning "JIRA" or "Confluence" in the summary	
$UnassignedIssues | Where-Object {$_.Summary -Match 'Confluence|JIRA'} | 
	Set-Ticket -Values $JIRA

Remove-JiraSession