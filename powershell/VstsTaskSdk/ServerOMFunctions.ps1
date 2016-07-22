<#
.SYNOPSIS
Gets a credentials object that can be used with the TFS extended client SDK.

.DESCRIPTION
The agent job token is used to construct the credentials object. The identity associated with the token depends on the scope selected in the build/release definition (either the project collection build/release service identity, or the project build/release service identity).

.EXAMPLE
$serverOMDirectory = Get-VstsTaskVariable -Name 'Agent.ServerOMDirectory' -Require
Add-Type -LiteralPath ([System.IO.Path]::Combine($serverOMDirectory, 'Microsoft.TeamFoundation.Client.dll'))
Add-Type -LiteralPath ([System.IO.Path]::Combine($serverOMDirectory, 'Microsoft.TeamFoundation.Common.dll'))
Add-Type -LiteralPath ([System.IO.Path]::Combine($serverOMDirectory, 'Microsoft.TeamFoundation.VersionControl.Client.dll'))
$tfsTeamProjectCollection = New-Object Microsoft.TeamFoundation.Client.TfsTeamProjectCollection(
    (Get-VstsTaskVariable -Name 'System.TeamFoundationCollectionUri' -Require),
    (Get-VstsTfsClientCredentials))
$versionControlServer = $tfsTeamProjectCollection.GetService([Microsoft.TeamFoundation.VersionControl.Client.VersionControlServer])
$versionControlServer.GetItems('$/*').Items | Format-List
#>
function Get-TfsClientCredentials {
    [CmdletBinding()]
    param([string]$OMDirectory)

    Trace-EnteringInvocation -InvocationInfo $MyInvocation
    try {
        $ErrorActionPreference = 'Stop'

        # Default the OMDirectory to the directory containing the entry script.
        if (!$OMDirectory) {
            $OMDirectory = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\..")
        }

        # Get the endpoint.
        $endpoint = Get-Endpoint -Name SystemVssConnection -Require

        # Testing the type will load the fallback DLL if required.
        $fallbackDll = [System.IO.Path]::Combine($OMDirectory, 'Microsoft.TeamFoundation.Client.dll')
        $typeName = 'Microsoft.TeamFoundation.Client.TfsClientCredentials'
        if (!(Test-Type -TypeName $typeName -FallbackDll $fallbackDll)) {
            # Bubble the error.
            $null = [type]$typeName
        }

        # Construct the credentials.
        $credentials = New-Object Microsoft.TeamFoundation.Client.TfsClientCredentials($false) # Do not use default credentials.
        $credentials.AllowInteractive = $false
        $credentials.Federated = New-Object Microsoft.TeamFoundation.Client.OAuthTokenCredential([string]$endpoint.auth.parameters.AccessToken)
        return $credentials
    } catch {
        $ErrorActionPreference = 'Continue'
        Write-Error $_
    } finally {
        Trace-LeavingInvocation -InvocationInfo $MyInvocation
    }
}

<#
.SYNOPSIS
Gets a credentials object that can be used with the REST SDK.

.DESCRIPTION
The agent job token is used to construct the credentials object. The identity associated with the token depends on the scope selected in the build/release definition (either the project collection build/release service identity, or the project service build/release identity).

.EXAMPLE
$vssCredentials = Get-VstsVssCredentials
Add-Type -LiteralPath (Assert-Path "$(Get-TaskVariable -Name 'Agent.ServerOMDirectory' -Require)\Microsoft.TeamFoundation.Common.dll" -PassThru)
Add-Type -LiteralPath (Assert-Path "$(Get-TaskVariable -Name 'Agent.ServerOMDirectory' -Require)\Microsoft.VisualStudio.Services.WebApi.dll" -PassThru)
# This is a bad example. All of the Server OM DLLs should be in one folder.
Add-Type -LiteralPath (Assert-Path "$(Get-TaskVariable -Name 'Agent.ServerOMDirectory' -Require)\Modules\Microsoft.TeamFoundation.DistributedTask.Task.Internal\Microsoft.TeamFoundation.Core.WebApi.dll" -PassThru)
$projectHttpClient = New-Object Microsoft.TeamFoundation.Core.WebApi.ProjectHttpClient(
    (New-Object System.Uri($env:SYSTEM_TEAMFOUNDATIONCOLLECTIONURI)),
    $vssCredentials)
$projectHttpClient.GetProjects().Result
#>
function Get-VssCredentials {
    [CmdletBinding()]
    param([string]$OMDirectory)

    Trace-EnteringInvocation -InvocationInfo $MyInvocation
    try {
        $ErrorActionPreference = 'Stop'

        # Default the OMDirectory to the directory containing the entry script.
        if (!$OMDirectory) {
            $OMDirectory = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\..")
        }

        # Get the endpoint.
        $endpoint = Get-Endpoint -Name SystemVssConnection -Require

        # Check if the VssOAuthAccessTokenCredential type is available.
        $fallbackDll = [System.IO.Path]::Combine($OMDirectory, 'Microsoft.VisualStudio.Services.WebApi.dll')
        if ((Test-Type -TypeName Microsoft.VisualStudio.Services.OAuth.VssOAuthAccessTokenCredential -FallbackDll $fallbackDll)) {
            # Return the credentials.
            return New-Object Microsoft.VisualStudio.Services.Common.VssCredentials(
                (New-Object Microsoft.VisualStudio.Services.Common.WindowsCredential($false)), # Do not use default credentials.
                (New-Object Microsoft.VisualStudio.Services.OAuth.VssOAuthAccessTokenCredential($endpoint.auth.parameters.AccessToken)),
                [Microsoft.VisualStudio.Services.Common.CredentialPromptType]::DoNotPrompt)
        }

        throw 'not impl'

        # Add-Type -LiteralPath (Assert-Path "$(Get-TaskVariable -Name 'Agent.ServerOMDirectory' -Require)\Microsoft.VisualStudio.Services.Common.dll" -PassThru)
        # Add-Type -LiteralPath (Assert-Path "$(Get-TaskVariable -Name 'Agent.ServerOMDirectory' -Require)\Microsoft.VisualStudio.Services.WebApi.dll" -PassThru)
        # # Add-Type -LiteralPath (Assert-Path "$(Get-TaskVariable -Name 'Agent.ServerOMDirectory' -Require)\Microsoft.VisualStudio.Services.Client.dll" -PassThru)
        # New-Object Microsoft.VisualStudio.Services.Common.VssCredentials(
        #     (New-Object Microsoft.VisualStudio.Services.Common.WindowsCredential($false)), # Do not use default credentials.
        #     (New-Object Microsoft.VisualStudio.Services.OAuth.VssOAuthAccessTokenCredential($endpoint.auth.parameters.AccessToken)),
        #     [Microsoft.VisualStudio.Services.Common.CredentialPromptType]::DoNotPrompt)

        # Add-Type -LiteralPath (Assert-Path "$(Get-TaskVariable -Name 'Agent.ServerOMDirectory' -Require)\Microsoft.VisualStudio.Services.Common.dll" -PassThru)
        # $endpoint = (Get-Endpoint -Name SystemVssConnection -Require)
        # New-Object Microsoft.VisualStudio.Services.Common.VssServiceIdentityCredential(
        #     New-Object Microsoft.VisualStudio.Services.Common.VssServiceIdentityToken([string]$endpoint.auth.parameters.AccessToken))
    } catch {
        $ErrorActionPreference = 'Continue'
        Write-Error $_
    } finally {
        Trace-LeavingInvocation -InvocationInfo $MyInvocation
    }
}

function New-VssHttpClient {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TypeName,

        [string]$OMDirectory,

        [string]$Uri,

        $VssCredentials)

    Trace-EnteringInvocation -InvocationInfo $MyInvocation
    try {
        $ErrorActionPreference = 'Stop'

        # Default the OMDirectory to the directory containing the entry script.
        if (!$OMDirectory) {
            $OMDirectory = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\..")
        }

        # Default the URI to the collection URI.
        if (!$Uri) {
            $Uri = Get-TaskVariable -Name System.TeamFoundationCollectionUri -Require
        }

        # Cast the URI.
        [uri]$Uri = New-Object System.Uri($Uri)

        # Default the credentials.
        if (!$VssCredentials) {
            $VssCredentials = Get-VssCredentials -OMDirectory $OMDirectory
        }

        # Determine the fallback DLL in case the type fails to load.
        $fallbackDll = $null
        if ($TypeName -like 'Microsoft.*.WebApi.*HttpClient') {
            # Try to interpret the fallback DLL name. Trim the last ".___HttpClient" and add ".dll".
            $fallbackDll = $TypeName.SubString(0, $TypeName.LastIndexOf('.'))
            $fallbackDll = [System.IO.Path]::Combine($OMDirectory, "$fallbackDll.dll")
        }

        # Test the type.
        if (!(Test-Type -TypeName $TypeName -FallbackDll $fallbackDll)) {
            # Bubble the error.
            $null = [type]$TypeName
        }

        # Try to construct the HTTP client.
        Write-Verbose "Constructing HTTP client."
        try {
            New-Object $TypeName($Uri, $VssCredentials)
        } catch {
            # Check if the exception is not due to Newtonsoft.Json DLL not found.
            if ($_.Exception.InnerException -isnot [System.IO.FileNotFoundException] -or
                $_.Exception.InnerException.FileName -notlike 'Newtonsoft.Json, *') {

                # Rethrow.
                throw
            }

            # Test if the Newtonsoft.Json DLL exists in the OM directory.
            $newtonsoftDll = [System.IO.Path]::Combine($OMDirectory, "Newtonsoft.Json.dll")
            Write-Verbose "Testing file path: '$newtonsoftDll'"
            if (!(Test-Path -LiteralPath $newtonsoftDll -PathType Leaf)) {
                Write-Verbose 'Not found. Rethrowing exception.'
                throw
            }

            # Add a binding redirect and try again. Parts of the Dev15 preview SDK have a
            # dependency on the 6.0.0.0 Newtonsoft.Json DLL, while other parts reference
            # the 8.0.0.0 Newtonsoft.Json DLL.
            Write-Verbose "Adding assembly resolver."
            $onAssemblyResolve = [System.ResolveEventHandler]{
                param($sender, $e)

                if ($e.Name -like 'Newtonsoft.Json, *') {
                    Write-Verbose "Resolving '$($e.Name)'"
                    return [System.Reflection.Assembly]::LoadFrom($newtonsoftDll)
                }

                Write-Verbose "Unable to resolve assembly name '$($e.Name)'"
                return $null
            }
            [System.AppDomain]::CurrentDomain.add_AssemblyResolve($onAssemblyResolve)
            try {
                # Try again to construct the HTTP client.
                Write-Verbose "Trying again to construct the HTTP client."
                New-Object $TypeName($Uri, $VssCredentials)
            } finally {
                # Unregister the assembly resolver.
                Write-Verbose "Removing assemlby resolver."
                [System.AppDomain]::CurrentDomain.remove_AssemblyResolve($onAssemblyResolve)
            }
        }
    } catch {
        $ErrorActionPreference = 'Continue'
        Write-Error $_
    } finally {
        Trace-LeavingInvocation -InvocationInfo $MyInvocation
    }
}

########################################
# Private functions.
########################################
function Test-Type {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TypeName,

        [string]$FallbackDll)

    Trace-EnteringInvocation -InvocationInfo $MyInvocation
    try {
        # Try to load the type.
        $ErrorActionPreference = 'Ignore'
        try {
            # Failure when attempting to cast a string to a type, transfers control to the
            # catch handler even when the error action preference is ignore. The error action
            # is set to Ignore so the $Error variable is not polluted.
            $null = [type]$TypeName

            # Success.
            return $true
        } catch { }

        $ErrorActionPreference = 'Stop'

        # Test if the fallback DLL exists.
        if ($FallbackDll) {
            Write-Verbose "Testing file path: '$FallbackDll'"
            if ((Test-Path -LiteralPath $FallbackDll -PathType Leaf)) {
                # Load the fallback DLL.
                Write-Verbose "Loading assembly: '$FallbackDll'"
                Add-Type -LiteralPath $FallbackDll

                # Try again.
                $ErrorActionPreference = 'Ignore'
                try {
                    # Failure when attempting to cast a string to a type, transfers control to the
                    # catch handler even when the error action preference is ignore. The error action
                    # is set to Ignore so the $Error variable is not polluted.
                    $null = [type]$TypeName

                    # Success.
                    return $true
                } catch { }

                $ErrorActionPreference = 'Stop'
            } else {
                # The fallback DLL was not found.
                Write-Verbose 'Not found.'
            }
        }

        return $false
    } finally {
        Trace-LeavingInvocation -InvocationInfo $MyInvocation
    }
}
