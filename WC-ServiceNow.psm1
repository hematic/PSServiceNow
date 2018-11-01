Function Get-WCServiceNowIncident{
    <#
        .SYNOPSIS
        Function to retrieve one or more ServiceNow Incident
        .DESCRIPTION
        This function allows you query the ServiceNow REST API using various filters to retrieve one or more incidents.
        .EXAMPLE
        $Splat = @{
            Credential = $Credential
            IncidentID = 'INC0010165'
            baseURI = "company.service-now.com"
        }
        Get-WCServiceNowIncident @Splat #Retrieve incident by incident ID
        .EXAMPLE
        $Splat = @{
            Credential = $Credential
            Description = 'Test'
            baseURI = "company.service-now.com"
        }
        Get-WCServiceNowIncident @Splat #Retrieve incident(s) by a short description filter.
        .EXAMPLE
        $Splat = @{
            Credential = $Credential
            User = '1a059879db9c2340873162eb0b961993'
            baseURI = "company.service-now.com"
        }
        Get-WCServiceNowIncident @Splat #Retrieve incident(s) filtering by ServiceNow sys user id.
        .EXAMPLE
        $Splat = @{
            Credential = $Credential
            User = 'username'
            baseURI = "company.service-now.com"
        }
        $BySamaccountname = Get-WCServiceNowIncident @Splat #Retrieve incident(s) filtering by samaccountname
        .EXAMPLE
        $Splat = @{
            Credential = $Credential
            User = 'first.last@company.com'
            baseURI = "company.service-now.com"
        }
        $ByEmail = Get-WCServiceNowIncident @Splat #Retrieve incident(s) filtering by email
        .EXAMPLE
        $Splat = @{
            Credential = $Credential
            User = 'Last, First'
            baseURI = "company.service-now.com"
        }
        $ByFirstandLast = Get-WCServiceNowIncident @Splat #Retrieve incident(s) filtering by First and last name
    #>
    [CmdletBinding(SupportsShouldProcess=$False,ConfirmImpact='Low',DefaultParameterSetName='ByID')]

    Param(
        [Parameter(Mandatory=$True,HelpMessage='This is credential for connecting to Service Now.')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$Credential,

        [Parameter(Mandatory=$True,HelpMessage='This is base URI of your Service Now instance.')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$baseURI,
        
        [Parameter(Mandatory=$True,ParameterSetName = 'ByIncidentID', HelpMessage='This is Ticket ID to search for.')]
        [ValidateNotNullOrEmpty()]
        [String]$IncidentID,
        
        [Parameter(Mandatory=$True,ParameterSetName = 'ByShortDescription', HelpMessage='This is part of the short description field to search for.')]
        [ValidateNotNullOrEmpty()]
        [String]$Description,
        
        [Parameter(Mandatory=$True,ParameterSetName = 'ByUser', HelpMessage='You can pass either the servicenow Userid, the email, the samaccountname, or the last, first names separated by commas.')]
        [ValidateNotNullOrEmpty()]
        [String]$User
    )

    #region Get Headers
    Try{
        $Headers = New-WCServiceNowHeaders -Credential $Credential -ErrorAction Stop
    }
    Catch{
        Write-Error $_
    }
    #endregion

    #region Build URI
    switch ($PsCmdlet.ParameterSetName)
    {
        'ByIncidentID' {
            $URI = "https://$BaseURI/api/now/table/incident?sysparm_query=number%3D$IncidentID&sysparm_limit=1"
        }

        'ByShortDescription' {
            $URI = "https://$baseURI/api/now/table/incident?sysparm_query=short_descriptionLIKE$Description"
        }
        
        'ByUser' {
            If($User -like '*@*'){
                Try{
                    $SysID = (Get-WCServiceNowUser -Credential $Credential -Email $User -ErrorAction Stop).'sys_id'
                }
                Catch{
                    Write-Error $_
                }
                
            }
            ElseIf($User -like '*,*'){
                $Splits = $User -split ','
                $Last = $Splits[0].TrimStart().TrimEnd()
                $First = $Splits[1].TrimStart().TrimEnd()
                Try{
                    $SysID = (Get-WCServiceNowUser -Credential $Credential -Firstname $First -LastName $Last).'sys_id'
                }
                Catch{
                    Write-Error $_
                }
            }
            ElseIf(Get-ADuser $User){
                $SysID = (Get-WCServiceNowUser -Credential $Credential -SamaccountName $User).'sys_id'
            }
            Else{
                $SysID = $User
            }
            
            $URI = "https://$baseURI/api/now/table/incident?sysparm_query=caller_id%3D$SysID"
        }
    }

    #endregion

    #region Send HTTP request
    Try{
        $Response = Invoke-WebRequest -Headers $Headers -Method Get -Uri $URI -ErrorAction Stop
        $Obj = ($Response.content | convertFrom-JSON).result
        Return $Obj
    }
    Catch{
        Write-error $_
    }
    #endregion
}
Function Get-WCServiceNowUser{
    <#
        .SYNOPSIS
        Function to retrieve a ServiceNowUser
        .DESCRIPTION
        This function is useful for retrieving the otherwise hidden sys_id of a user in ServiceNow.
        .EXAMPLE
        $Splat = @{
            Credential = $Credential
            SamaccountName = 'username'
            baseURI = "company.service-now.com"
        }
        Get-WCServiceNowUser @Splat #Retrieve user by samaccountname
        .EXAMPLE
        $Splat = @{
            Credential = $Credential
            Email = 'first.last@company.com'
            baseURI = "company.service-now.com"
        }
        Get-WCServiceNowUser @Splat #Retrieve user by email
        .EXAMPLE
        $Splat = @{
            Credential = $Credential
            First      = 'First'
            Last       = 'Last'
            baseURI = "company.service-now.com"
        }
        Get-WCServiceNowUser @Splat #Retrieve user by first and last name
    #>
    [CmdletBinding(SupportsShouldProcess=$False,ConfirmImpact='Low',DefaultParameterSetName='BySamAccountname')]
    Param(
        [Parameter(Mandatory=$True,HelpMessage='This is credential for connecting to Service Now.')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$Credential,

        [Parameter(Mandatory=$True,HelpMessage='This is base URI of your Service Now instance.')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$baseURI,
        
        [Parameter(Mandatory=$True,ParameterSetName = 'BySamAccountname', HelpMessage='This is Ticket ID to search for.')]
        [ValidateNotNullOrEmpty()]
        [String]$SamaccountName,
        
        [Parameter(Mandatory=$True,ParameterSetName = 'ByEmail', HelpMessage='This is part of the short description field to search for.')]
        [ValidateNotNullOrEmpty()]
        [String]$Email,
        
        [Parameter(Mandatory=$True,ParameterSetName = 'ByFirstandLast', HelpMessage='This is the user ID in ServiceNow.')]
        [ValidateNotNullOrEmpty()]
        [String]$Firstname,

        [Parameter(Mandatory=$True,ParameterSetName = 'ByFirstandLast', HelpMessage='This is the user ID in ServiceNow.')]
        [ValidateNotNullOrEmpty()]
        [String]$LastName
    )

    #region Get Headers
    Try{
        $Headers = New-WCServiceNowHeaders -Credential $Credential -ErrorAction Stop
    }
    Catch{
        Write-Error $_
    }
    #endregion

    #region Build URI
    switch ($PsCmdlet.ParameterSetName)
    {
        'BySamAccountname' {
            $URI = "https://$baseURI/api/now/table/sys_user?sysparm_query=user_name%3D$SamaccountName"
        }

        'ByEmail' {
            $URIEmail = $Email -replace '@', '%40'
            $URI = "https://$baseURI/api/now/table/sys_user?sysparm_query=email%3D$URIEmail"
        }
        
        'ByFirstandLast' {
            $URI = "https://$baseURI/api/now/table/sys_user?sysparm_query=first_name%3D$FirstName%5Elast_name%3D$Lastname"
        }
    }

    #endregion

    #region Send HTTP request
    Try{
        $Response = Invoke-WebRequest -Headers $Headers -Method Get -Uri $URI -ErrorAction Stop
        $Obj = ($Response.content | convertFrom-JSON).result
        Return $Obj
    }
    Catch{
        Write-error $_
    }
    #endregion
}
Function New-WCServiceNowIncident{
    <#
        .SYNOPSIS
        Function to create a new ServiceNow Incident
        .DESCRIPTION
        This function allows you pass as few or as many of the parameters as you would like to create a
        ServiceNow Incident. All are available, none are mandatory.
        .EXAMPLE
        $Splat = @{
            Credential        = $Credential
            baseURI           = "company.service-now.com"
            category          = 'Hardware'
            subcategory       = 'Desktop'
            u_symptom         = 'Crash'
            location          = 'Tampa'
            contact_type      = 'Email'
            impact            = 1
            urgency           = 1
            priority          = 1
            assignment_group  = 'Service Desk'
            caller_id         = 'username'
            short_description = 'This is the short description'
            comments          = 'This is a test comment'
            work_notes        = 'These are work notes'
        }
        New-WCServiceNowIncident @Splat #New Incident creation with all currently available fields
    #>
    [CmdletBinding(SupportsShouldProcess=$False,ConfirmImpact='Low')]
    Param(
        [Parameter(Mandatory=$True, HelpMessage='This is the API Credential.')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$Credential,

        [Parameter(Mandatory=$True,HelpMessage='This is base URI of your Service Now instance.')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$baseURI,

        [Parameter(Mandatory=$False, HelpMessage='This is the incident category.')]
        [ValidateSet('hardware','software_service')]
        [String]$category,

        [Parameter(Mandatory=$False, HelpMessage='This is the subcategory.')]
        [String]$subcategory,

        [Parameter(Mandatory=$False, HelpMessage='This is the symptom.')]
        [ValidateSet('slow_performance','error_message','crash','unable_to_launch_connect','access_issue')]
        [String]$u_symptom,

        [Parameter(Mandatory=$False, HelpMessage='This is the office location.')]
        [String]$location,

        [Parameter(Mandatory=$False, HelpMessage='This is the contact type.')]
        [ValidateSet('email','call','im','walk-in','non_user_query','self-service')]
        [String]$contact_type,

        [Parameter(Mandatory=$False, HelpMessage='This is the impact.')]
        [ValidateSet('1','2','3')]
        [Int]$impact,

        [Parameter(Mandatory=$False, HelpMessage='This is the urgency.')]
        [ValidateSet('1','2','3')]
        [Int]$urgency,

        [Parameter(Mandatory=$False, HelpMessage='This is the priority.')]
        [ValidateSet('1','2','3','4')]
        [Int]$priority,

        [Parameter(Mandatory=$False, HelpMessage='This is the assignment group.')]
        [String]$assignment_group,

        [Parameter(Mandatory=$False, HelpMessage='This is the affected user.')]
        [String]$caller_id,

        [Parameter(Mandatory=$False, HelpMessage='This is the short description which acts like a title.')]
        [String]$short_description,

        [Parameter(Mandatory=$False, HelpMessage='This is the assigned to user (not yet implemented).')]
        [String]$asssigned_to,

        [Parameter(Mandatory=$False, HelpMessage='This is a switch to suppress emails on a ticket.')]
        [Switch]$u_suppress_emails,

        [Parameter(Mandatory=$False, HelpMessage='This is the comment field that is publicly visible.')]
        [String]$comments,

        [Parameter(Mandatory=$False, HelpMessage='This is the work notes field only visible to technicians and admins.')]
        [String]$work_notes,

        [Parameter(Mandatory=$False, HelpMessage='This is the user the ticket is created on behalf of.')]
        [String]$u_on_behalf_of,

        [Parameter(Mandatory=$False, HelpMessage='This is the user the ticket is assigned to.')]
        [String]$assigned_to
    )

    #region Get Headers
    Try{
        $headers = New-WCServiceNowHeaders -Credential $Credential -ErrorAction Stop
    }
    Catch{
        Write-Error $_
    }
    #endregion
    
    #region Form Body
    $Body = @{}

    Foreach($Item in $PSBoundParameters.keys | Where-object {$_ -ne 'Credential'}){
        $Body.$Item = $PSBoundParameters.Item($Item)
    }
    
    $jsonbody = ConvertTo-Json $Body
    Write-host $JSONBody
    #endregion

    #region Send HTTP request
    Try{
        $URI = "https://$baseURI/api/now/table/incident"
        $Response = Invoke-WebRequest -Headers $Headers -Method Post -Uri $URI -Body $JSONBody -ContentType 'application/json' -ErrorAction Stop
        $Obj = ($Response.content | convertFrom-JSON).result
        Return $Obj
    }
    Catch{
        Write-error $_
    }
    #endregion
}
Function Set-WCServiceNowIncident{
    <#
        .SYNOPSIS
        Function to update a new ServiceNow Incident
        .DESCRIPTION
        This function allows you pass as few or as many of the parameters as you would like to update a
        ServiceNow Incident. All are available, none are mandatory.
        .EXAMPLE
        $Splat = @{
            Credential        = $Credential
            baseURI           = "company.service-now.com"
            IncidentID        = 'INC0010165'
            impact            = 2
            caller_id         = 'username'
            short_description = 'This is the updated short description'
            comments          = 'This is a second test comment'

        }
        Set-WCServiceNowIncident @Splat
    #>
    [CmdletBinding(SupportsShouldProcess=$False,ConfirmImpact='Low')]
    Param(
        [Parameter(Mandatory=$True, HelpMessage='This is the API Credential.')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$Credential,

        [Parameter(Mandatory=$True,HelpMessage='This is base URI of your Service Now instance.')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$baseURI,

        [Parameter(Mandatory=$True, HelpMessage='This is the API Credential.')]
        [ValidateNotNullOrEmpty()]
        [String]$IncidentID,

        [Parameter(Mandatory=$False, HelpMessage='This is the incident category.')]
        [ValidateSet('hardware','software_service')]
        [String]$category,

        [Parameter(Mandatory=$False, HelpMessage='This is the subcategory.')]
        [String]$subcategory,

        [Parameter(Mandatory=$False, HelpMessage='This is the symptom.')]
        [ValidateSet('slow_performance','error_message','crash','unable_to_launch_connect','access_issue')]
        [String]$u_symptom,

        [Parameter(Mandatory=$False, HelpMessage='This is the office location.')]
        [String]$location,

        [Parameter(Mandatory=$False, HelpMessage='This is the contact type.')]
        [ValidateSet('email','call','im','walk-in','non_user_query','self-service')]
        [String]$contact_type,

        [Parameter(Mandatory=$False, HelpMessage='This is the impact.')]
        [ValidateSet('1','2','3')]
        [Int]$impact,

        [Parameter(Mandatory=$False, HelpMessage='This is the urgency.')]
        [ValidateSet('1','2','3')]
        [Int]$urgency,

        [Parameter(Mandatory=$False, HelpMessage='This is the priority.')]
        [ValidateSet('1','2','3','4')]
        [Int]$priority,

        [Parameter(Mandatory=$False, HelpMessage='This is the assignment group.')]
        [String]$assignment_group,

        [Parameter(Mandatory=$False, HelpMessage='This is the affected user.')]
        [String]$caller_id,

        [Parameter(Mandatory=$False, HelpMessage='This is the short description which acts like a title.')]
        [String]$short_description,

        [Parameter(Mandatory=$False, HelpMessage='This is the assigned to user (not yet implemented).')]
        [String]$asssigned_to,

        [Parameter(Mandatory=$False, HelpMessage='This is a switch to suppress emails on a ticket.')]
        [Switch]$u_suppress_emails,

        [Parameter(Mandatory=$False, HelpMessage='This is the comment field that is publicly visible.')]
        [String]$comments,

        [Parameter(Mandatory=$False, HelpMessage='This is the work notes field only visible to technicians and admins.')]
        [String]$work_notes,

        [Parameter(Mandatory=$False, HelpMessage='This is the user the ticket is created on behalf of.')]
        [String]$u_on_behalf_of,

        [Parameter(Mandatory=$False, HelpMessage='This is the user the ticket is assigned to.')]
        [String]$assigned_to,

        [Parameter(Mandatory=$False, HelpMessage='This is the status of the ticket.')]
        [ValidateSet('1','2','3','6','7','8')]
        [Int]$incident_state
    )

    #region Get Headers
    Try{
        $Headers = New-WCServiceNowHeaders -Credential $Credential -ErrorAction Stop
    }
    Catch{
        Write-Error $_
    }
    #endregion

    #region Get SysID
    Try{
        $SysID = (Get-WCServiceNowIncident -Credential $Credential -IncidentID $IncidentID -ErrorAction Stop).'Sys_ID'
    }
    Catch{
        Write-Error $_
    }
    #endregion

    #region Build URI
    $URI = "https://$baseURI/api/now/table/incident/$SysID"
    #endregion

    #region Build Body
    $Body = @{}

    Foreach($Item in $PSBoundParameters.keys | Where-object {$_ -ne 'Credential' -and $_ -ne 'IncidentID'}){
        $Body.$Item = $PSBoundParameters.Item($Item)
    }
    
    $jsonbody = ConvertTo-Json $Body

    #endregion

    #region Send HTTP request
    Try{
        $Response = Invoke-WebRequest -Headers $Headers -Method Put -Uri $URI -Body $JSONBody -ContentType 'application/json' -ErrorAction Stop
        $Obj = ($Response.content | convertFrom-JSON).result
        Return $Obj
    }
    Catch{
        Write-error $_
    }
    #endregion
}
Function New-WCServiceNowBase64AuthObj{
    [CmdletBinding(SupportsShouldProcess=$False,ConfirmImpact='Low')]
    Param(
        [Parameter(Mandatory=$True, HelpMessage='This is the API Credential.')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$Credential
    )
    [String]$textuser = $Credential.UserName
    [String]$textPassword = $Credential.GetNetworkCredential().password
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $TextUser, $Textpassword)))
    Return $base64AuthInfo
}
Function New-WCServiceNowHeaders{
    [CmdletBinding(SupportsShouldProcess=$False,ConfirmImpact='Low')]
    Param(
        [Parameter(Mandatory=$True, HelpMessage='This is the API Credential.')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$Credential
    )
    Try{
        $base64AuthInfo = New-WCServiceNowBase64AuthObj -Credential $Credential -ErrorAction Stop
    }
    Catch{
        Write-Error $_
    }
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
    $headers.Add('Accept','application/json')
    Return $headers
}