####################################################################
#ConnectWise Retrieval Functions

Function Get-CWTicket
{
	<#
	.SYNOPSIS
		Retrieves an object formatted ConnectWise Ticket.
	
	.DESCRIPTION
		Pass this function a ticket id and if the ticket exists it
        will return you all the information. If it doesnt exist it 
        will return $False.
	
	.Parameter Ticketid
        This is the ConnectWise ticket ID that you want to retrive information on.

	.EXAMPLE
		Get-CWTicket -TicketID '8675309'

	.EXAMPLE
		Get-CWTicket '8675309'
    #>

    [cmdletbinding()]
    
    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
		[INT]$TicketID
    )

    Begin
    {
    [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/service/tickets/$ticketID"
    [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
    [string]$ContentType = "application/json"
    [string]$Authstring  = $CWInfo.company + '+' + $CWCredentials.publickey + ':' + $CWCredentials.privatekey
    $encodedAuth         = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($Authstring)));

    $Headers=@{
        'X-cw-overridessl' = "True"
        "Authorization"="Basic $encodedAuth"
        }
     }
    
    Process
    {      
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers  -ContentType $ContentType -Method Get
    }
    
    End
    {
        If($JSONResponse)
        {
            Return $JSONResponse
        }

        Else
        {
            Return $False
        }
    }
}

Function Get-CWTimeEntries
{
	<#
	.SYNOPSIS
		Retrieves an array of Time Entries related to a ConnectWise Ticket.
	
	.DESCRIPTION
		Pass this function a ticket id and if the ticket exists and their 
        are time entries it will return you all the information. If it 
        doesnt exist it will return $False.
	
	.Parameter Ticketid
        This is the ConnectWise ticket ID that you want to retrive time
        entries for.

	.EXAMPLE
		Get-CWTimeEntries -TicketID '8675309'

	.EXAMPLE
		Get-CWTimeEntries '8675309'
    #>
    
    [cmdletbinding()]

    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[INT]$TicketID
    )

    Begin
    {
    [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/service/tickets/$ticketID/timeentries"
    [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
    [string]$ContentType = "application/json"
    [string]$Authstring  = $CWInfo.company + '+' + $CWCredentials.publickey + ':' + $CWCredentials.privatekey
    $encodedAuth         = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($Authstring)));

    $Headers=@{
        'X-cw-overridessl' = "True"
        "Authorization"="Basic $encodedAuth"
        }
     }
    
    Process
    {   
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers  -ContentType $ContentType -Method Get
    }

    End
    {
        If($JSONResponse)
                {
        Return $JSONResponse
    }

        Else
                {
        Return $False
    }
    }
}

Function Get-TimeEntryDetails
{
	<#
	.SYNOPSIS
		Retrieves an object formatted time entry record.
	
	.DESCRIPTION
		Pass this function a time entry id and if the entry exists it
        will return you all the information. If it doesnt exist it 
        will return $False. Time Entry ID's can be gathered from the
        Get-CWTimeEntries function.
	
	.Parameter TimeentryID
        This is the ConnectWise time entry ID that you want to retrive
        specific data for.

 	.EXAMPLE
		Get-TimeEntryDetails -$TimeEntryID '7714622'

	.EXAMPLE
		Get-TimeEntryDetails '7714622'
    #>

    [cmdletbinding()]

    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[INT]$TimeEntryID
    )

    Begin
    {
    [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/Time/Entries/$TimeEntryID"
    [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
    [string]$ContentType = "application/json"
    [string]$Authstring  = $CWInfo.company + '+' + $CWCredentials.publickey + ':' + $CWCredentials.privatekey
    $encodedAuth         = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($Authstring)));

    $Headers=@{
        'X-cw-overridessl' = "True"
        "Authorization"="Basic $encodedAuth"
        }
    }
    
    Process
    {   
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers  -ContentType $ContentType -Method Get
    }

    End
    {
        If($JSONResponse)
                {
        Return $JSONResponse
    }

        Else
                {
        Return $False
    }
    }
}

Function Get-CWMember
{

    <#
	.SYNOPSIS
		Retrieves an object formatted ConnectWise Member.
	
	.DESCRIPTION
		Pass this function a email address and if it matches a member
        it will return you all the information. If it doesnt exist it 
        will return $False.

    .Parameter EmailAddress
        This is the email address belonging to the member you are 
        trying to retrieve information on.
	
	.EXAMPLE
		Get-CWMember -EmailAddress 'pmarshall@labtechsoftware.com'

	.EXAMPLE
		Get-CWMember 'pmarshall@labtechsoftware.com'
    #>

    [cmdletbinding()]

    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$EmailAddress
    )

    Begin
    {
        [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/system/members"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"
        [string]$Authstring  = $CWInfo.company + '+' + $CWCredentials.publickey + ':' + $CWCredentials.privatekey
        $encodedAuth         = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($Authstring)));

                    $Headers=@{
        'X-cw-overridessl' = "True"
        "Authorization"="Basic $encodedAuth"
        }
    }

    Process
    {
        $Body = @{
    "conditions" = "emailaddress = '$EmailAddress'"
    
    }
       
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -ContentType $ContentType -Body $Body -Method Get
    }

    End
    {
        If($JSONResponse)
                {
        Return $JSONResponse
    }

        Else
                {
        Return $False
    }
    }
}

Function Get-CWContact
{

    <#
	.SYNOPSIS
		Retrieves an object formatted ConnectWise Contact.
	
	.DESCRIPTION
		Pass this function a first name,last name and company ID
        if there is a contact it will return you the ID. If it 
        doesn't exist it will return $False.
	
    .Parameter First
        The First Name of the contact you are searching for.

    .Parameter Last
        The Last Name of the contact you are searching for.

    .Parameter CompanyID
        The companyid that the contact belongs to.
        
	.EXAMPLE
		Get-CWContact -First 'Phillip' -Last 'Marshall' -CompanyID '49804'

	.EXAMPLE
		Get-CWContact 'Phillip' 'Marshall' '49804'

    #>

    [cmdletbinding()]

    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$First,
    	[Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$Last,
    	[Parameter(Mandatory = $true,Position = 2,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$CompanyID
    )

    Begin
    {
        [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/company/contacts"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"
        [string]$Authstring  = $CWInfo.company + '+' + $CWCredentials.publickey + ':' + $CWCredentials.privatekey
        $encodedAuth         = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($Authstring)));

                    $Headers=@{
        'X-cw-overridessl' = "True"
        "Authorization"="Basic $encodedAuth"
        }
    }

    Process
    {
        $Body = @{
    "conditions" = "firstname = '$First' AND lastname = '$Last' AND company/id =$CompanyID"
    }
      
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -ContentType $ContentType -Body $Body -Method Get
    }

    End
    {
        If($JSONResponse)
        {
        Return $JSONResponse.id
    }

        Else
        {
        Return $False
    }
    }
}

function Get-CWKeys
{
	<#
	.SYNOPSIS
		Authenticates with the ConnectWise API.
	
	.DESCRIPTION
		This function builds the authstring and makes an authentication
        call to impersonate the member you have chosen. Returns you access keys.
	
	.Parameter ImpersonationMember
        This is the Member you are impersonating (by name).

	.EXAMPLE
		Get-CWKeys -ImpersonationMember 'Jira'

	.EXAMPLE
		Get-CWKeys 'Jira'
    #>

    [cmdletbinding()]
    
    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
		[String]$ImpersonationMember
    )

    Begin
    {
        [string]$BaseUri     = "$CWServerRoot" + "v4_6_release/apis/3.0/system/members/$ImpersonationMember/tokens"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"
        [string]$Authstring  = $CWInfo.company + '+' + $CWInfo.user + ':' + $CWInfo.password

        #Convert the user and pass (aka public and private key) to base64 encoding
        $encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($Authstring)));

        #Create Header
        $header = [Hashtable] @{
            Authorization = ("Basic {0}" -f $encodedAuth)
            Accept        = $Accept
            Type          = "application/json"
            'X-cw-overridessl' = "True"
            'x-cw-usertype' = "integrator" 
        };
    }

    Process
    {
        $body   = "{`"memberIdentifier`":`"$ImpersonationMember`"}"
    
        #execute the the request
        $response = Invoke-RestMethod -Uri $Baseuri -Method Post -Headers $header -Body $body -ContentType $contentType;
    }

    End
    {
        Return $response;
    }
}     

Function Get-CWBoardInfo
{
	<#
	.SYNOPSIS
		Retrieves an object formatted ConnectWise Board.
	
	.DESCRIPTION
		This function will return all boards based on the filter conditions
        in the body. You should supply some conditions as there may be too
        many boards to return otherwise.
	
	.Parameter BoardName
        The name of the board of the partial name of the board you are
        trying to retrieve information on.
	
	.EXAMPLE
		Get-CWBoardInfo -BoardName $Boardname

	.EXAMPLE
		Get-CWBoardInfo $Boardname
    #>

    [cmdletbinding()]

    param
    (
        [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [String]$BoardName
	)
    
    Begin
    {
        [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/service/boards"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"
        [string]$Authstring  = $CWInfo.company + '+' + $CWCredentials.publickey + ':' + $CWCredentials.privatekey
        $encodedAuth         = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($Authstring)));

        $Headers=@{
        'X-cw-overridessl' = "True"
        "Authorization"="Basic $encodedAuth"
        }
    }

    Process
    {
        $Body = @{
        "fields" = "id,name"
        "conditions" = "name LIKE `'%$Boardname%`'"
        }

       
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -ContentType $ContentType -Body $Body -Method Get
    }
    
    End
    {   
        If($JSONResponse)
        {
            Return $JSONResponse
        }

        Else
        {
            Return $False
        }
    }
}

####################################################################
#ConnectWise Post/Edit Functions

function New-CWTicket
{
	<#
	.SYNOPSIS
		Creates a new ticket on a CW Board.
	
	.DESCRIPTION
		You pass this function a ticket object and it will create
        a ticket on the proper board.
	
    .DETAILED
        The Ticket Object should contain at minimum:
        $Ticket.Summary
        $Ticket.assignee.emailaddress

    .Parameter Ticket
        The ticket object containing information needed to create 
        the new ticket.

    .Parameter Boardname
        The boardname to create the ticket on.
        
	.EXAMPLE
		New-CWTicket -Ticket $Issue

	.EXAMPLE
		New-CWTicket $Issue
    #>

    [cmdletbinding()]

    param
    (
        [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [Object]$Ticket,
        [Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [String]$BoardName
    )

    Begin
    {
        [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/service/tickets"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"
        [string]$Authstring  = $CWInfo.company + '+' + $CWCredentials.publickey + ':' + $CWCredentials.privatekey
        $encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($Authstring)));
    }

    Process
    {
        #Making sure the summary field is formatted properly
        ###############################################################
        If($($Ticket.Summary.length) -gt 90)
        {
        $Summary = $($Ticket.Summary.substring(0,75))
    }

        $Summary = Format-SanitizedString -InputString $($Ticket.Summary)
        $Summary = $Summary.Replace('"', "'")
        ###############################################################
    
        $UserInfo = Get-ProperUserInfo -JiraEmail $($Ticket.assignee.emailaddress)
        $Body= @"
{
    "summary"   :    "[JIRA][$($Ticket.key)] - $($Summary)",
    "board"     :    {"name": "$BoardName"},
    "status"    :    {"name": "New"},
    "company"   :    {"id": "$($UserInfo.CompanyID)"},
    "contact"   :    {"id": "$($UserInfo.ContactID)"}
}
"@
        $Headers=@{
'X-cw-overridessl' = "True"
"Authorization"="Basic $encodedAuth"
}
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -ContentType $ContentType -Method Post -Body $Body
    }

    End
    {
        If($JSONResponse)
        {
        Return $JSONResponse
    }

        Else
        {
        Return $False
    }
    }
}

function New-CWTimeEntry
{

	<#
	.SYNOPSIS
		Makes a time Entry for a user
	
	.DESCRIPTION
		This function will make a time entry in ConnectWise
        for whatever user you specify. You must Pass it a 
        Worklog Object.
	
	.DETAILED
        The worklog object should contain at a minimum:
            $Worklog.ID
            $Worklog.CWTicketID.
            $Worklog.created
            $Worklog.ended
            $Worklog.comment
	
    .PARAMETER Worklog
        This is a formatted [PSCustomObject] that contains the worklog
        information from Jira.
    
    .PARAMETER CWTicketID
        This is the Connectwise Ticket ID you want to make a time
        entry on.

	.EXAMPLE
		New-CWTimeEntry -Worklog $Worklog -CWTicketID '8675309'

	.EXAMPLE
		New-CWTimeEntry "$Worklog" "8675309"
    #>

    [cmdletbinding()]

    param
    (
        [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [PSCustomObject]$WorkLog,
        [Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [Int]$CWTicketID
    )

    Begin
    {
        [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/time/entries"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"
        [string]$Authstring  = $CWInfo.company + '+' + $CWCredentials.publickey + ':' + $CWCredentials.privatekey
        $encodedAuth         = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($Authstring)));
    }

    Process
    {
        #Date Magic
        $StartedUniversal = (get-date ($WorkLog.started)).ToUniversalTime()
        [string]$Ended = ($StartedUniversal).AddSeconds($Worklog.timeSpentSeconds)
        [String]$Created = Get-Date ($StartedUniversal) -format "yyyy-MM-ddTHH:mm:ssZ"
        [String]$Ended = Get-Date ($Ended) -format "yyyy-MM-ddTHH:mm:ssZ"

        #Member Magic
        $MemberInfo = Get-ProperUserInfo -JiraEmail "$($Worklog.author.emailaddress)" -MemberCheck '1'

        $Body= @"
{
    "Summary"        : "$($Worklog.id)",
    "chargeToType"   : "ServiceTicket",
    "chargeToId"     : "$CWTicketID",
    "timeStart"      : "$Created",
    "timeend"        : "$Ended",
    "internalnotes"  : "[JiraID!!$($Worklog.id)!!] $($Worklog.comment)",
    "company"        : {"id": "$($Memberinfo.CompanyID)"},
    "member"         : {"id": "$($Memberinfo.MemberID)"},
    "billableOption" : "DoNotBill"
    
}
"@
        $Headers=@{
'X-cw-overridessl' = "True"
"Authorization"="Basic $encodedAuth"
}
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -ContentType $ContentType -Method Post -Body $Body
    }

    End
    {
        Return $JSONResponse
    }
}

function Invoke-TicketProcess
{
	<#
	.SYNOPSIS
		Processes a ticket between Jira and ConnectWise
	
	.DESCRIPTION
		This function will take a ticket in Jira and check if the
        value in the CW Ticket custom field is filled out. If it is
        it will check if that ticket exists in CW. If needed a new 
        ticket will be made and the custom field will be updated with 
        the proper value.
	
	.DETAILED
	
    .PARAMETER Issue
        This is a formatted [PSCustomObject] that contains the ticket
        information from Jira.

    .PARAMETER BoardName
        This is the ConnectWise board you want tickets created on.
    
	.EXAMPLE
		Invoke-TicketProcess -Issue $Issue

	.EXAMPLE
		Invoke-TicketProcess $Issue
    #>

    [cmdletbinding()]

    param
    (
        [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [PSObject]$Issue,
        [Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [String]$Boardname
    )

    Process
    {
        If ($Issue.CWTicketID -eq 'None')
        {
            $Ticket = New-CWTicket -Ticket $Issue -BoardName "$BoardName"
            $issue.CWTicketID = $Ticket.id
            #Update-CWTicketValue -CWTicketID $($Ticket.id) -JiraIssueKey $($Issue.Key)
            Edit-JiraIssue -IssueID "$($Issue.ID)" -CWTicketID "$($Ticket.id)"
        }

        Else
        {
            $CurrentTicket = Get-cwticket -TicketID $Issue.cwticketid

            If($CurrentTicket.id -ne $Issue.cwticketid)
            {
              $Ticket = New-CWTicket -Ticket $Issue -BoardName "LT-Infrastructure"
              Edit-JiraIssue -IssueID "$($Issue.ID)" -CWTicketID "$($Ticket.id)"
              $issue.CWTicketID = $Ticket.id  
            }
        }
    }
}

function Invoke-WorklogProcess
{
	<#
	.SYNOPSIS
		Processes a time entry between Jira and ConnectWise
	
	.DESCRIPTION
		This function will take a timeentry in Jira and check if the
        time entry exists on the CW Ticket. If it doesn't it will create it.
	
	.DETAILED
	
    .PARAMETER Issue
        This is a formatted [PSCustomObject] that contains the ticket
        information from Jira.
    
	.EXAMPLE
		Invoke-WorklogProcess -Issue $Issue

	.EXAMPLE
		Invoke-WorklogProcess $Issue
    #>

    [cmdletbinding()]

    param
    (
        [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [Object]$Issue
    )

    Process
    {
        [Array]$TimeEntryIDs = (Get-CWTimeEntries -TicketID "$($Issue.cwticketid)").id

        Foreach($TimeEntry in $TimeEntryIDs)
        {
            [Array]$TEDetails += Get-TimeEntryDetails -TimeEntryID $TimeEntry
        }
  
        Foreach($Worklog in $Issue.WorkLog.worklogs)
        {
            Foreach($Detail in $TEDetails)
            {
                [INT]$Present = 0
                $RegCheck = ([regex]::matches($Detail.internalnotes, "(?:\[JiraID!!)(.*)(?:!!)")).groups[1].value
                
                If($Worklog.id -eq $RegCheck)
                {
                    [INT]$Present = 1
                }  
        }

            If($Present -ne 1)
            {
                New-CWTimeEntry -WorkLog $Worklog -CWTicketID "$($Issue.cwticketid)"
            }
            #A Section needs to be added here to verify that the worklog entry hasn't changed.
            #I'll come back to this.
    }
    }
}