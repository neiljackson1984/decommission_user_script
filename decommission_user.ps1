#!pwsh

# this script sets up the appropriate (non-trivial) mail transport rules (and potentially inbox rules) 
# necessary to achieve the following behavior:
# Any message sent to primaryEmailAddressOfUserToBeDecommissioned ends up in the mailbox of primaryEmailAddressOfUserToBeDecommissioned
# and in the mailboxes specified by redirectDestinations.  The sender will receive a bounceback message.

[CmdletBinding()]
Param (

    [String]$primaryEmailAddressOfUserToBeDecommissioned,
    
    [Parameter(HelpMessage=
        @"
We will create (or at least ensure the existence of) a mailbox having 
this alias as its primary alias.  We will call the primary email address of this mailbox the rejectionTargetAddress.
The existence of this mailbox ensures that we can have a transport rule that redirects to rejectionTargetAddress
without triggering a "could-not-find address" bounceback.
"@
    )]
    [String]$rejectionTargetAlias = 'rejection',
    
    [Parameter(HelpMessage=
        "This is a list of the email addresses to which we want to redirect messages sent to primaryEmailAddressOfUserToBeDecommissioned"
    )]
    [String[]]$redirectDestinations = @(),
    
    [Parameter(HelpMessage=
        "list of names of users that are to be granted full access to the mailbox that is being decommissioned"
    )]
    [String[]]$usersToBeGrantedFullAccess = @(),
    
    [Parameter(HelpMessage=
        "The bounceback message will contain a message advising the sender to re-send his message to this email alias."
    )]
    [String]$emailAliasToAdviseSendersToSendTo  = 'info',

    [Parameter(HelpMessage=
        "The path of the powershell script to dot source in order to ensure that we are connected to the appropriate cloud administration modules."
    )]
    [String]$pathOfScriptToConnectToCloud,
    
    [Parameter(HelpMessage=
        "specifies whether to set the automapping property on the mailbox permissions that we assign to usersToBeGrantedFullAccess"
    )]
    [Boolean]$automapping=$false
    
    
)

Begin {
    Write-Host "Now decommissioning user $primaryEmailAddressOfUserToBeDecommissioned"
    
    #dot source the script to connect to the appropriate cloud administration modules
    . $pathOfScriptToConnectToCloud
}

Process {

    $emailAddressToBeDecommissioned = $primaryEmailAddressOfUserToBeDecommissioned
    $primaryDomainName = (Get-AzureAdDomain | where {$_.IsDefault}).Name
    $ruleNamePrefix= $emailAddressToBeDecommissioned + '--decommission'
    $emailAddressToAdviseSendersToSendTo = "$emailAliasToAdviseSendersToSendTo@$primaryDomainName"
    $rejectionTargetAddress = "$rejectionTargetAlias@$primaryDomainName"
    
    $magicHeader1Name = 'X-' + [guid]::NewGuid().ToString().Replace('-','')
    $magicHeader1Value = ''  + [guid]::NewGuid().ToString().Replace('-','')
    $magicHeader2Name = 'X-' + [guid]::NewGuid().ToString().Replace('-','')
    $magicHeader2Value = ''  + [guid]::NewGuid().ToString().Replace('-','')
        
    # confirm that the rejection target address exists:
    if (-not (Get-Recipient $rejectionTargetAddress -ErrorAction SilentlyContinue)){
        New-Mailbox -Shared -Name $rejectionTargetAddress -DisplayName $rejectionTargetAddress -Alias ($rejectionTargetAddress -split '@')[0]
    }
    
    # I am equivocating on whether to implement the redirects with inbox rule or transport rule.
    # the potential benefit of using an inbox rule instead of a transport rule is that the user to whom messages are being redirected
    # could, assuming he has full access permissions to the decommiossioned mailbox, delete the inbox rule himself, whereas the same user might
    # need to invoke the help of an administrator in order to delete the transport rule.
    # the disadvantage of the inbox rule is that if a redirectionDestination is unreachable, a bounceback will show up in the decommissioned mailbox (which would be confusing 
    # to the decommissioned user if he ever was recommissioned.)
    # $redirectDestinationsToBeHandledByTransportRule = @();    $redirectDestinationsToBeHandledByInboxRule     = $redirectDestinations[0..($redirectDestinations.Count - 1)]
    $redirectDestinationsToBeHandledByInboxRule     = @();    $redirectDestinationsToBeHandledByTransportRule = $redirectDestinations[0..($redirectDestinations.Count - 1)]
    
    
    $augmentedRedirectDestinations = $redirectDestinationsToBeHandledByTransportRule + $primaryEmailAddressOfUserToBeDecommissioned
    
    
    $i = 1
    $recipes = @(

        { New-TransportRule -Name ( $ruleNamePrefix +  ($script:i++) ) `
            -ExceptIfHeaderContainsMessageHeader $magicHeader1Name  -ExceptIfHeaderContainsWords $magicHeader1Value `
            -SentTo $emailAddressToBeDecommissioned `
            -RecipientAddressType Resolved `
            -RedirectMessageTo $rejectionTargetAddress `
            -BlindCopyTo $augmentedRedirectDestinations `
            -StopRuleProcessing $False `
            -SetHeaderName $magicHeader1Name -SetHeaderValue $magicHeader1Value `
            # -RouteMessageOutboundConnector 'loopback' `
            # -HeaderContainsMessageHeader $magicHeader1Name  -HeaderContainsWords $magicHeader1Value `
            # -PrependSubject 'a' `
        },        
        
                                        
        { New-TransportRule -Name ( $ruleNamePrefix +  ($script:i++) ) `
            -HeaderContainsMessageHeader $magicHeader1Name  -HeaderContainsWords $magicHeader1Value `
            -SentTo $rejectionTargetAddress `
            -RecipientAddressType Resolved `
            -RedirectMessageTo $emailAddressToBeDecommissioned `
            -SetHeaderName $magicHeader2Name -SetHeaderValue $magicHeader2Value `
            -StopRuleProcessing $False `
        },
        
        # $rejectionTargetAddress must actually exist as a recipient, else a could-not-find bounceback is generated immediately upon attmepting redirect.
        
        # We need two independently turn-on-able, independently test-able flags (in this case the magic headers 1 and 2, respectively),
        # and at least flag 1 (or the combination of flag 1 and the rejection target address) need to be unique to the emailAddressToBeDecommissioned.
        # it probably makes since to have a single generic "rejection" mailbox and to let flag1 vary (one way that would make sense would be to have a fixed 
        # name and a value that depended on emailAddressToBeDecommissioned.
                                        
        { New-TransportRule -Name ( $ruleNamePrefix +  ($script:i++) ) `
            -HeaderContainsMessageHeader $magicHeader2Name  -HeaderContainsWords $magicHeader2Value `
            -SentTo $emailAddressToBeDecommissioned `
            -RecipientAddressType Resolved `
            -RejectMessageEnhancedStatusCode 5.7.1  -RejectMessageReasonText "Your message to $emailAddressToBeDecommissioned was refused.  Please re-send your message to $emailAddressToAdviseSendersToSendTo ." `
            -StopRuleProcessing $False `
        }
        
        # in the case where we have two decommisioned email addresses, and one is a redirect destination for the other, you can have, depending on which address's decommissioning transport rules come first, 
        # a sender to one of the addresses can get two bounceback messages.  We probably should create and test for some guard headers to prevent this problem.

        
    )

    $transportRuleNamePattern = '^' + $ruleNamePrefix + '\d*' + '$'

    # remove any transport rules left over from previous runs of this script:
    $existingTransportRulesToRemove = @(Get-TransportRule | where-object {$_.Name -match $transportRuleNamePattern} )
    
    Write-Host "Found $($existingTransportRulesToRemove.Count) existing transport rules that we will now remove."
    $existingTransportRulesToRemove | Remove-TransportRule -Confirm:$False
    
    
    
    #execute the recipes
    $recipes | foreach-object {& $_}
    
    #spit out the newly-created transport rules:
    Get-TransportRule | where-object {$_.Name -match $transportRuleNamePattern} | select Name, Priority
    
    
    $inboxRuleNamePattern = '^' + $ruleNamePrefix + '\d+' + '_.*' + '$'
    $existingInboxRulesToRemove = @(Get-InboxRule -Mailbox $emailAddressToBeDecommissioned | where-object {$_.Name -match $inboxRuleNamePattern} )
    Write-Host "Found $($existingInboxRulesToRemove.Count) existing inbox rules that we will now remove."
    $existingInboxRulesToRemove | Remove-InboxRule -Confirm:$False
    $i=1
    $newlyCreatedInboxRules = @(
        foreach ( $redirectDestination in $redirectDestinationsToBeHandledByInboxRule ){
            New-InboxRule `
                -Confirm:$false  `
                -Mailbox $emailAddressToBeDecommissioned   `
                -Name ( $ruleNamePrefix + ($i++) + "_" + "redirect_to_" + $redirectDestination )     `
                -HeaderContainsWords $magicHeader1Value `
                -RedirectTo $redirectDestination  `
                -StopProcessingRules:$false `
        }
    )
    Write-Host "newlyCreatedInboxRules: "
    $newlyCreatedInboxRules
    
    foreach ( $userToBeGrantedFullAccess in $usersToBeGrantedFullAccess){
        #remove any existing full access permission so that we can re-set the permission with our preferred value for Automapping
        Remove-MailboxPermission -Identity $emailAddressToBeDecommissioned -User    $userToBeGrantedFullAccess -AccessRights FullAccess -Confirm:$false -ErrorAction SilentlyContinue
        
        Add-MailboxPermission   -Identity $emailAddressToBeDecommissioned -User    $userToBeGrantedFullAccess -AccessRights FullAccess -Automapping $automapping 
        
        Add-RecipientPermission -Identity $emailAddressToBeDecommissioned -Trustee $userToBeGrantedFullAccess -AccessRights SendAs  -confirm:$false
    }
    
}

End {
    
}
