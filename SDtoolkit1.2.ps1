# Version 1.20
### Done
# All AAD CMDlets were removed, AAD powershell sdk was also removed
### To be done
# Try functioning the commands from a differnet script eventually

$Menu = {
    Write-Host " *********************************"
    Write-Host " *           SD Toolkit          *"
    Write-Host " *********************************"
    Write-Host " * 1.) PS SDK Connections        *"
    Write-Host " * 2.) Managers                  *"
    Write-Host " * 3.) User Management           *"
    Write-Host " * 4.)                           *"
    Write-Host " * 5.)                           *"
    Write-Host " * 6.)                           *"
    Write-Host " * 7.) Quit                      *"
    Write-Host " *********************************"
    Write-Host " Selection: " -nonewline
    }
    
    
    Do 
    {
    Clear-Host
    Invoke-Command $Menu
    $SelectMainMenu = Read-Host
        Switch ($SelectMainMenu) 
        {
            1 # Connections
                {
    
                $MenuConnections = {
                Write-Host " *********************************"
                Write-Host " *         Connections           *"
                Write-Host " *********************************"
                Write-Host " * 1.) MgGraph                   *"
                Write-Host " * 2.) ExchangeOnline            *"
                Write-Host " * 3.) PNPOnline                 *"
                Write-Host " * 4.) Main Menu                 *"
                Write-Host " *********************************"
                Write-Host " Selection: " -nonewline
                }
    
                    Do 
                    {
                    Clear-Host
                    Invoke-Command $MenuConnections
                    $SelectConnections = Read-Host
                    Clear-Host
                        Switch ($SelectConnections) 
                        {
                            1  # MgGraph
                                {
                                Do 
                                {
                                Clear-Host
                                Write-Host '1. Connect-MgGraph'
                                Write-Host '2. Disconnect-MgGraph'
                                Write-Host '3. Main Menu'
                                $SelectMgGraph = Read-Host
                                Clear-Host
                                     Switch ($SelectMgGraph)
                                     {
                                        1
                                            {
                                            Connect-MgGraph 
                                            Pause
                                            }
                                        2
                                            {
                                            Disconnect-MgGraph
                                            Pause
                                            }
                                     }
                                }
                                While ($SelectMgGraph -ne 3)
                                }
                            2  # ExchangeOnline 
                                {
                                Connect-ExchangeOnline
                                }
                            3  # PNPOnline
                                {
                                    Do  
                                    {
                                    Clear-Host
                                    Write-Host '1. SME Bank'
                                    Write-Host '2. SME Finance'
                                    Write-Host '3. Main Menu'
                                    $SelectPNPOnline = Read-Host
                                    Clear-Host
                                         Switch ($SelectPNPOnline)
                                         {
                                            1  
                                                {
                                                $TenantAdminURLSMEB = "https://smedigitalfinance-admin.sharepoint.com/"
                                                Connect-PnPOnline -Url $TenantAdminURLSMEB -interactive
                                                }
                                            2
                                                {
                                                $TenantAdminURLSMEF = "https://uabsmefinance-admin.sharepoint.com/"
                                                Connect-PnPOnline -Url $TenantAdminURLSMEF -interactive
                                                pause
                                                }
                                         }
                                    }
                                    While ($SelectPNPOnline -ne 3)
                                } 
                            TestConnections   
                                {
                                Write-Host "MGGraph:" 
                                (Get-MgUser -Filter "startswith(userprincipalname,'Admin_D@')").userprincipalname
                                Write-Host "Exchange Online:"    
                                Get-ConnectionInformation | Format-List state, userprincipalname
                                Write-Host "PNPonline:"
                                Get-PnPConnection | Format-List Connectiontype, url
                                pause
                                }  
                        }
                    }
                    While ($SelectConnections -ne 4)
                }
            2 # Managers
                {
    
                $MenuManagers = {
                Write-Host " *********************************"
                Write-Host " *           Managers            *"
                Write-Host " *********************************"
                Write-Host " * 1.) Direct Reports            *" # MgGraph
                Write-Host " * 2.) Replace Team Manager      *" # MgGraph
                Write-Host " * 3.) Temporary Team Manager    *" # MgGraph
                Write-Host " * 4.) Permenant Team Manager    *" # MgGraph
                Write-Host " * 5.) Description               *"
                Write-Host " * 6.) Main Menu                 *"
                Write-Host " *********************************"
                Write-Host " Selection: " -nonewline
                }
    
                    Do 
                    {
                    Clear-Host
                    Invoke-Command $MenuManagers
                    $SelectManagers = Read-Host
                    Clear-Host
                        Switch ($SelectManagers) 
                        {
                            1 # Direct Reports 
                                {
                                Clear-Host
                                # Custom Variables
                                $ManagerUPN = Read-Host -Prompt 'Input Manager UPN'
                                $manager = (Get-MgUser -userid $ManagerUPN).id
    
                                # Get Direct reports
                                $directReports = Get-MgUserDirectReport -UserId $manager
                                $directReports | ForEach-Object { Get-MgUser -userid $_.Id } | Select-Object -ExpandProperty UserPrincipalName
                                Pause
                                }
                            2 # Replace Team Manager
                                {
                                Clear-Host
                                # Custom Variables
                                $ManagerUPN = Read-Host -Prompt 'Current Team Manager'
                                $ReplacementManagerUPN = Read-Host -Prompt 'New Team Manager'
                                $manager = (Get-MgUser -userid $ManagerUPN).id
                                $ReplacementManager = (Get-MgUser -userid $ReplacementManagerUPN).id
                                $ReplacementManagerMG = @{ "@odata.id"="https://graph.microsoft.com/v1.0/users/$ReplacementManager" }
    
                                # Replace Manager
                                $directReports = Get-MgUserDirectReport -UserId $manager
                                foreach ($employee in $directReports) 
                                     {
                                        Set-MgUserManagerByRef -UserId $($employee.id) -BodyParameter $ReplacementManagerMG
                                     }
                                Write-Host "$replacementManagerUPN team members:"
                                Get-MgUserDirectReport -UserId $ReplacementManager |  ForEach-Object { Get-MgUser -userid $_.Id } | Select-Object -ExpandProperty UserPrincipalName
                                pause
                                }
                            3 # Temporary Team Manager
                                {
                                Clear-Host
                                # Custom Variables
                                $ManagerUPN = Read-Host -Prompt 'Current Team Manager UPN'
                                $TempManagerUPN = Read-Host -Prompt 'Input Temporary Manager UPN'
                                $Manager = (Get-MgUser -userid $ManagerUPN).id
                                $TempManager = (Get-MgUser -userid $TempManagerUPN).id
                                $TempGroupName = $ManagerUPN.Split('@')[0]
                                $TempManagerMG = @{ "@odata.id"="https://graph.microsoft.com/v1.0/users/$TempManager" }
    
                                New-MgGroup -DisplayName "TempManager-$TempGroupName" -MailEnabled:$false -SecurityEnabled:$true -MailNickName "TempManager-$TempGroupName" 
                                $TempGroupNameID = (Get-MgGroup -Filter "DisplayName eq 'TempManager-$TempGroupName'").id
                                $directReports = Get-MgUserDirectReport -UserId $manager
                                foreach ($employee in $directReports) 
                                     {
                                     New-MgGroupMember -GroupId $TempGroupNameID -DirectoryObjectId $($employee.Id)
                                     Set-MgUserManagerByRef -UserId $($employee.id) -BodyParameter $TempManagerMG
                                     }
                                pause
                                }
                            4 # Permenant Team Manager
                                {
                                Clear-Host
                                # Custom Variables
                                $Groups = Get-MgGroup -Filter "startswith(displayName, 'tempmanager-')" | Select-Object -ExpandProperty Displayname
                                $Groups # Display current tempmanager groups
                                $TempManagerGroup = Read-Host -Prompt 'Select the Group that the users will be moved from'
                                $PermManagerUPN = Read-Host -Prompt 'New Manager UPN'
                                $PermManager = (Get-MgUser -userid $PermManagerUPN).id
                                $PermManagerMG = @{ "@odata.id"="https://graph.microsoft.com/v1.0/users/$PermManager" }
    
                                $TempGroupNameID = (Get-MgGroup -Filter "DisplayName eq '$TempManagerGroup'").id
                                $TempManagerGroupMembers = Get-MgGroupMember -GroupId $TempGroupNameID
                                foreach ($employee in $TempManagerGroupMembers) 
                                    {
                                    Set-MgUserManagerByRef -UserId $($employee.id) -BodyParameter $PermManagerMG
                                    }
                                    Remove-MgGroup -GroupId $TempGroupNameID
                                Pause
                                }
                            5 # Description  
                                {
                                Clear-Host
                                Write-Host "1. Provides all the direct reports associated with the provided managers email"
                                Write-Host "2. Takes all the current team members from the chosen manager and moves the to the selected user"
                                Write-Host "3. Takes all the current team members from the chosen manager, creates a custom EntraID group, asignes the users to it and moves the members to temporary manager of choice"
                                Write-Host "4. Displays custom groups made by the 3rd option, after choosing one of the groups, it takes all the members from it, assigns a newly selected manager to them, and deletes the group"
                                Pause
                                }
                        }
                    }
                    While ($SelectManagers -ne 6)
                }
            3 # User Management
                {
                $MenuUserManagmemet = {
                Write-Host " *********************************"
                Write-Host " *         User Management       *"
                Write-Host " *********************************"
                Write-Host " * 1.) Disabling user [full]     *" # MgGraph ExchangeOnline PNPonline
                Write-Host " * 2.)                           *"
                Write-Host " * 3.)                           *"
                Write-Host " * 4.)                           *"
                Write-Host " * 5.) Main Menu                 *"
                Write-Host " *********************************"
                Write-Host " Selection: " -nonewline
                }
    
                Do 
                    {
                    Clear-Host
                    Invoke-Command $MenuUserManagmemet
                    $SelectUserManagement = Read-Host
                    Clear-Host
                        Switch ($SelectUserManagement) 
                        {
                            1  # Disabling user [full]
                                {
                                    # Custom variables
                                    $UserUPN = Read-Host -Prompt 'Input user UPN'
                                    $User = (Get-MgUser -userid $UserUPN).id
                                    $UserPNP = "i:0#.f|membership|$UserUPN"
                                    $RecipientType = (get-mailbox -identity $UserUPN).RecipientTypeDetails
                                    
                                    # Convert mailbox to shared
                                        if ($RecipientType -eq "UserMailbox")
                                            {
                                            set-mailbox -identity $userUPN -type shared
                                            Write-Host "$UserUPN converted to SharedMailbox" -ForegroundColor Green
                                            }
                                        else 
                                            {
                                            Write-Host "$UserUPN is already a SharedMailbox" -ForegroundColor Yellow
                                            }

                                    # Remove manager
                                        Remove-MgUserManagerByRef -UserId $User
                                        Write-Host "$UserUPN manager removed" -ForegroundColor Green
    
                                    # Remove user from all Groups
                                        $Groups = (Get-MgUserMemberOf -userid $user).id
                                        foreach ($Group in $Groups)
                                            {
                                                $securityenabled = (get-mggroup -groupid $group).securityenabled
                                                $mailenabled = (get-mggroup -groupid $group).mailenabled
                                                $Grouptypes = (get-mggroup -groupid $group).Grouptypes
                                                If (($mailenabled -eq $False) -and ($securityenabled -eq $True) -and ([string]::IsNullOrEmpty($Grouptypes))) # Security groups
                                                {Remove-MgGroupMemberDirectoryObjectByRef -GroupId $Group -DirectoryObjectId $user}
                                                If (($mailenabled -eq $True) -and ($securityenabled -eq $False) -and ([string]::IsNullOrEmpty($Grouptypes))) # DL
                                                {Remove-DistributionGroupMember -identity $group -member $user -BypassSecurityGroupManagerCheck -Confirm:$false}
                                                If (($mailenabled -eq $True) -and ($securityenabled -eq $True) -and ([string]::IsNullOrEmpty($Grouptypes))) # Mail enabled
                                                {Remove-DistributionGroupMember -identity $group -member $user -BypassSecurityGroupManagerCheck -Confirm:$false}
                                                If (($mailenabled -eq $True) -and ($securityenabled -eq $False) -and ($Grouptypes -eq "unified")) # Microsoft 365
                                                {Remove-MgGroupMemberDirectoryObjectByRef -GroupId $Group -DirectoryObjectId $user}
                                            }
                                        Write-Host "All group memberships were removed" -ForegroundColor Green

                                    # Remove office licenses
                                        $UserLicenses = (Get-MgUserLicenseDetail -UserID $UserUPN).SkuId
                                        Set-MgUserLicense -UserId $UserUPN -RemoveLicenses $UserLicenses -AddLicenses @()
                                        Write-Host "Office licenses were removed" -ForegroundColor Green
        
                                    # Remove accesses from sharepoint
                                        Write-Host "Removing all Sharepoint accesses..."
                                        $TempURL = (Get-PnPTenantSite | Where-Object {$_.Url -like "*-my.sharepoint.com/*"}).url
                                        $ProperURL = $tempUrl.Replace("-my","")
                                        #Get All Site collections - Filter BOT and MySite Host
                                        $Sites = Get-PnPTenantSite -Filter "Url -like '$ProperURL'"
                                        #Iterate through all sites
                                        $Sites | ForEach-Object {
                                        try {
                                            Connect-PnPOnline -Url $_.URL -Interactive -ErrorAction SilentlyContinue
                                            if((Get-PnPUser | Where-Object {$_.LoginName -eq $UserPNP}) -ne $NULL) {
                                                try {
                                                    Remove-PnPUser -Identity $UserPNP -Confirm:$false -ErrorAction SilentlyContinue
                                                    } 
                                                catch {# Suppress error message
                                                }
                                            }
                                            } 
                                        catch {# Suppress error message
                                        }
                                        }
                                        Write-Host "$UserUPN removed from all sharepoint sites" -ForegroundColor Green
                                        
                                    # Block Sign in
                                        $params = @{AccountEnabled = "false"}
                                        Update-MgUser -UserId $User -BodyParameter $params 
                                        Write-Host "$UserUPN sign in was blocked" -ForegroundColor Green
                                        Pause
                                }
                            2
                                {
                                    
                                }
    
    
                        } 
                    } 
                    While ($SelectUserManagement -ne 5)
    
                }
        }    
    }
    While ($SelectMainMenu -ne 7)
