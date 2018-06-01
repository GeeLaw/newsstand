#Requires -Modules WebAuthenticationBroker

<#
.SYNOPSIS
    Sets an Inbox Rule to mark messages from certain senders with a certain
    category on Microsoft-powered e-mail accounts.

.DESCRIPTION
    This script works with Microsoft account and Azure AD account.

    The script reads one or more files in "newsstand" format that specifies
    e-mail addresses from which the messages are to be regularly swept. It
    then asks the user to sign in with their personal or organisation account.
    The script requires MailboxSettings.ReadWrite permission to inspect Inbox
    Rules and to create/update the Inbox Rule to automatically mark messages.
    The rule will be created as if it were created with Outlook Web Access.

    Specifically, before using the script, the user needs to create a "To be
    swept" category (the name can be different, but requires a change of
    parameter). By default, the script manages a rule whose display name is
    "Mark as To be swept (dcb300de84dc4a54a287659ed611977e)".

    To use the script, supply it with at least one listing file. Sign in with
    your Microsoft account or Azure AD account when prompted. Then consent the
    app for read/write access to your mailbox settings. After that, a blank
    page is shown. The URL will look like the following:

    https://login.microsoftonline.com/common/oauth2/nativeclient?code=...

    Copy it from the address bar and paste the complete URL into the prompt
    to continue.

.PARAMETER ListingFile
    Paths to listing files. The values can be piped to the command.

    Note that a listing file can include other files with "++ path/to/other"
    syntax. An included file need not be explicitly specified here.

.PARAMETER RecursionDepth
    The maximum depth of recursive including.

    Defaults to 16. Valid range for this parameter is [5, 64].

.PARAMETER Sequence
    The order of this rule.

    Defaults to 1. Valid range for this parameter is [1, 1024].

.PARAMETER RuleCategory
    The name of the category to be applied.

    Defaults to "To be swept".

.PARAMETER RuleDisplayName
    The display name of the rule (used for both look-up and creation).

    Defaults to "Mark as To be swept (dcb300de84dc4a54a287659ed611977e)".

.PARAMETER Offline
    If this switch is on, no online work will be done, and when parsing
    completes, an interactive window of summary is shown.

    Turning on this switch makes the script ignore -WhatIf and -Confirm.

.PARAMETER WhatIf
    If this switch is on, the script does not make changes to Outlook Web
    Access. After it has inspected current Inbox Rules, it displays what it
    decides to do: either create a new rule or update an exisiting rule.
    It also displays an interactive window of summary.

.PARAMETER Confirm
    If this switch is on, the script displays the modification it will make
    along with an interactive window of summary, and asks for confirmation
    from the user.

.EXAMPLE
    .\Update-MarkAsToBeSweptRule.ps1 list.txt

    This example parses list.txt and sets the rule accordingly.

.LINK
    https://github.com/GeeLaw/newsstand

#>
[CmdletBinding(SupportsShouldProcess = $true)]
Param
(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [string[]]$ListingFile,
    [Parameter(Mandatory = $false)]
    [ValidateRange(5, 64)]
    [UInt32]$RecursionDepth = 16,
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 1024)]
    [UInt32]$Sequence = 1,
    [Parameter(Mandatory = $false)]
    [string]$RuleCategory = 'To be swept',
    [Parameter(Mandatory = $false)]
    [string]$RuleDisplayName = 'Mark as To be swept (dcb300de84dc4a54a287659ed611977e)',
    [Parameter(Mandatory = $false)]
    [switch]$Offline,
    [Parameter(Mandatory = $false)]
    [switch]$WaitUI
)
Begin
{
    $recipientList = [System.Collections.Generic.List[object]]::new();
    $emailSet = [System.Collections.Generic.HashSet[string]]::new();
    $recurseOnFile = [System.Collections.Generic.List[string]]::new();
}
Process
{
    $recurseDepth = 0;
    While ($ListingFile.Length -gt 0)
    {
        If ((++$recurseDepth) -eq $RecursionDepth)
        {
            Write-Error -Message "Recursion is too deep (limit = $RecursionDepth)." `
                -ErrorId 'recurseTooDeep' `
                -Category LimitsExceeded `
                -RecommendedAction 'Use less deep recursion or increase recursion depth limit.';
            Break;
        }
        $recurseOnFile.Clear();
        ForEach ($file in (Resolve-Path $ListingFile).Path)
        {
            Write-Verbose "Processing file: $file";
            $currentRecipientName = $null;
            Get-Content -LiteralPath $file -Encoding UTF8 |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                ForEach-Object {
                    $_ = $_.Trim();
                    If ($_.StartsWith('%% '))
                    {
                        # A comment line.
                    }
                    If ($_.StartsWith('++ '))
                    {
                        # Include another file later.
                        $recurseOnFile.Add([System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($file), $_.Substring(3).TrimStart()));
                    }
                    ElseIf ($_.StartsWith('@@ '))
                    {
                        # Change current recipient name.
                        $currentRecipientName = $_.Substring(3).TrimStart();
                    }
                    ElseIf (-not $emailSet.Add($_.ToLowerInvariant()))
                    {
                        # The email is already present, skip it.
                        Write-Verbose "$_ is duplicated in the list.";
                    }
                    ElseIf ($currentRecipientName -eq $null)
                    {
                        # The name falls back to the address.
                        $local:rcp = New-Object PSObject -Property @{ 'name' = $_; 'address' = $_ };
                        $rcp = New-Object PSObject -Property @{ 'emailAddress' = $rcp };
                        $recipientList.Add($rcp);
                    }
                    Else
                    {
                        $local:rcp = New-Object PSObject -Property @{ 'name' = $currentRecipientName; 'address' = $_ };
                        $rcp = New-Object PSObject -Property @{ 'emailAddress' = $rcp };
                        $recipientList.Add($rcp);
                    }
                };
        }
        $ListingFile = $recurseOnFile.ToArray();
    }
    If ($RecursionDepth -eq $recurseDepth)
    {
        Break;
    }
}
End
{
    If ($recipientList.Count -eq 0)
    {
        Write-Error -Message 'No email addresses found.' `
            -ErrorId 'noAddressFound' `
            -Category ObjectNotFound `
            -RecommendedAction 'Include more files, or edit the files.';
        Break;
    }
    Write-Verbose "$($recipientList.Count) e-mail address(es) collected.";
    If ($Offline)
    {
        Write-Verbose 'Opening an interactive window for inspection.';
        $recipientList.emailAddress | Out-GridView -Title 'Newsstand Senders' -Wait:$WaitUI;
        Write-Verbose 'Working offline, skipping other steps.';
        Break;
    }
    Write-Verbose 'Requesting a token.';
    $accessToken = 0..0 | ForEach-Object {
        # Build the URI of the page for authorisation.
        $local:clientId = '95cab5f9-e6cb-4c83-842a-d7b4a3041397';
        $local:responseType = 'code';
        $local:redirectUri = 'https://login.microsoftonline.com/common/oauth2/nativeclient';
        $redirectUri = [uri]::EscapeDataString($redirectUri);
        $local:scope = 'https://graph.microsoft.com/mailboxsettings.readwrite';
        $scope = [uri]::EscapeDataString($scope);
        $local:uiUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
        $uiUrl += '?client_id=' + $clientId;
        $uiUrl += '&response_type=' + $responseType;
        $uiUrl += '&redirect_uri=' + $redirectUri;
        $uiUrl += '&scope=' + $scope;
        # Launch the page and wait for the result.
        $local:retUri = Request-WebAuthentication -InitialUri $uiUrl -CompletionExtractor {
            If ($_.ToLowerInvariant().StartsWith('https://login.microsoftonline.com/common/oauth2/nativeclient?'))
            {
                Return ($_.Substring('https://login.microsoftonline.com/common/oauth2/nativeclient?'.Length));
            }
        } -Title 'Sign in with your Microsoft account or Azure AD account';
        $retUri = '&' + $retUri + '&';
        # Check if there is an error.
        $local:errorRegex = [regex]::new('&[eE][rR][rR][oO][rR]=(.*?)&');
        $local:errorMatch = $errorRegex.Match($retUri);
        If ($errorMatch.Success)
        {
            $errorMatch = [uri]::UnescapeDataString($errorMatch.Groups[1].Value);
            Write-Error "Authentication failed, reason: $errorMatch." -ErrorId $errorMatch -Category AuthenticationError;
            Return;
        }
        # Look for the authorization code.
        $local:codeRegex = [regex]::new('&[cC][oO][dD][eE]=(.*?)&');
        $local:codeMatch = $codeRegex.Match($retUri);
        If (-not $codeMatch.Success)
        {
            Write-Error -Message 'The return URI does not contain a code.' -ErrorId 'no_code' -Category AuthenticationError;
            Return;
        }
        # Redeem access token.
        $local:grantType = 'authorization_code';
        $local:code = $codeMatch.Groups[1].Value;
        $local:codeResult = Invoke-RestMethod -UseBasicParsing `
            -Method POST -Uri 'https://login.microsoftonline.com/common/oauth2/v2.0/token' `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body "client_id=$clientId&redirect_uri=$redirectUri&scope=$scope&grant_type=$grantType&code=$code";
        If ($codeResult.error -ne $null)
        {
            Write-Error -Message $codeResult.error_description -Category AuthenticationError -ErrorId $codeResult.error;
            Return;
        }
        If ($codeResult.access_token -eq $null)
        {
            # An exception has been thrown during the request.
            Return;
        }
        $codeResult.access_token;
    };
    If ($accessToken -eq $null)
    {
        Break;
    }
    # Find current rules.
    $local:currentRule = Invoke-RestMethod -UseBasicParsing `
        -Method GET -Uri 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messagerules' `
        -Headers @{ 'Authorization' = "Bearer $accessToken" };
    If ($currentRule.error -ne $null)
    {
        Write-Error -Message "An error occurred while listing the rules: $($currentRule.error.message)" `
            -ErrorId $currentRule.error.code `
            -TargetObject $currentRule `
            -Category InvalidResult;
        Break;
    }
    If ([object]::ReferenceEquals($currentRule.value, $null))
    {
        # An exception has been thrown during the request.
        Break;
    }
    # Find the one with the specific name.
    $currentRule = @($currentRule.value | Where-Object displayName -eq $RuleDisplayName);
    If ($currentRule.Count -gt 1)
    {
        Write-Error -Message "There are at least two rules named `"$RuleDisplayName`"." `
            -ErrorId 'cannotResolveRuleFromName' `
            -TargetObject $currentRule `
            -Category OperationStoppped;
        Break;
    }
    # If there wasn't a matching rule, create our own;
    # otherwise, patch the current rule.
    # Since the name contains a newly generated GUID,
    # any sensible person will not use it as the display
    # name of any rule managed by himself or herself.
    $local:fakeRuleId = '?a99597c029de419596a0fa47bd6a8dc7'
    $local:currentRuleId = $fakeRuleId;
    If ($currentRule.Count -eq 0)
    {
        $currentRule = New-Object PSObject | Select-Object -Property actions, conditions, displayName, exceptions, isEnabled, isReadOnly, sequence;
        Write-Verbose 'No matching rule is found, will create a new rule.';
    }
    If ($currentRule.Count -eq 1)
    {
        $currentRule = $currentRule[0];
        $currentRuleId = $currentRule.id;;
        $currentRule = $currentRule | Select-Object -Property actions, conditions, displayName, exceptions, isEnabled, isReadOnly, sequence;
        Write-Verbose "Rule [$currentRuleId] is matched and will be updated.";
    }
    $currentRule.actions = New-Object PSObject -Property @{ 'assignCategories' = @($RuleCategory); 'stopProcessingRules' = $False };
    $currentRule.conditions = New-Object PSObject -Property @{ 'fromAddresses' = $recipientList };
    $currentRule.displayName = $RuleDisplayName;
    $currentRule.exceptions = New-Object PSObject;
    $currentRule.isEnabled = $true;
    $currentRule.isReadOnly = $false;
    $currentRule.sequence = $Sequence;
    # Prepare payload and send the request.
    $local:payload = $currentRule | ConvertTo-Json -Compress -Depth 32;
    $payload = [System.Text.Encoding]::UTF8.GetBytes($payload);
    $local:method = 'POST';
    $local:ruleUri = 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messagerules';
    $local:hintTitle = 'Creating a new rule';
    $local:hintFinish = 'Created a new rule [{0}].';
    If ($currentRuleId -ne $fakeRuleId)
    {
        $method = 'PATCH';
        $ruleUri = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messagerules('$currentRuleId')";
        $hintTitle = 'Updating rule [{0}]';
        $hintFinish = 'Updated rule [{0}].';
    }
    # We need to display UI before asking for confirmation,
    # therefore, a direct check is necessary.
    If ($WhatIfPreference -or [int]$ConfirmPreference -lt [int][System.Management.Automation.ConfirmImpact]::Medium)
    {
        Write-Verbose 'Opening an interactive window for inspection.';
        $recipientList.emailAddress | Out-GridView -Title 'Newsstand Senders' -Wait:$WaitUI;
        If (-not $PSCmdlet.ShouldProcess("Sending $method request to $ruleUri with the payload shown in the interactive window.",
            "Do you want to send $method request to $ruleUri with the payload shown in the interactive window?",
            [string]::Format($hintTitle, $currentRuleId)))
        {
            Break;
        }
    }
    $currentRule = Invoke-RestMethod -UseBasicParsing -Method $method -Uri $ruleUri `
        -Headers @{ 'Authorization' = "Bearer $accessToken" } -ContentType 'application/json; charset=utf8' `
        -Body $payload;
    If ($currentRule.error)
    {
        Write-Error -Message "An error occurred while editing the rule: $($currentRule.error.message)" `
            -ErrorId $currentRule.error.code `
            -TargetObject $currentRule `
            -Category InvalidResult;
        Break;
    }
    ElseIf ($currentRule.id -eq $null)
    {
        # An exception has been thrown during the request.
        Break;
    }
    [string]::Format($hintFinish, $currentRule.id) | Write-Verbose;
}
