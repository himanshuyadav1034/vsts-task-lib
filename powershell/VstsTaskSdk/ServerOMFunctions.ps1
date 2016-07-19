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
    param()

    Add-Type -LiteralPath (Assert-Path "$(Get-TaskVariable -Name 'Agent.ServerOMDirectory' -Require)\Microsoft.TeamFoundation.Client.dll" -PassThru)
    $endpoint = (Get-Endpoint -Name SystemVssConnection -Require)
    $credentials = New-Object Microsoft.TeamFoundation.Client.TfsClientCredentials($false) # Do not use default credentials.
    $credentials.AllowInteractive = $false
    $credentials.Federated = New-Object Microsoft.TeamFoundation.Client.OAuthTokenCredential([string]$endpoint.auth.parameters.AccessToken)
    $credentials
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
    param(
        [string]$OMDirectory)

    Trace-EnteringInvocation -InvocationInfo $MyInvocation
    try {
        if (!$OMDirectory) {
            # Fallback to the directory containing the entry script.
            $OMDirectory = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\..")
        }

        # Get the credentials.
        $endpoint = (Get-Endpoint -Name SystemVssConnection -Require)
        Add-Type -LiteralPath (Assert-Path "$(Get-TaskVariable -Name 'Agent.ServerOMDirectory' -Require)\Microsoft.VisualStudio.Services.Common.dll" -PassThru)
        Add-Type -LiteralPath (Assert-Path "$(Get-TaskVariable -Name 'Agent.ServerOMDirectory' -Require)\Microsoft.VisualStudio.Services.WebApi.dll" -PassThru)
        # Add-Type -LiteralPath (Assert-Path "$(Get-TaskVariable -Name 'Agent.ServerOMDirectory' -Require)\Microsoft.VisualStudio.Services.Client.dll" -PassThru)
        New-Object Microsoft.VisualStudio.Services.Common.VssCredentials(
            (New-Object Microsoft.VisualStudio.Services.Common.WindowsCredential($false)), # Do not use default credentials.
            (New-Object Microsoft.VisualStudio.Services.OAuth.VssOAuthAccessTokenCredential($endpoint.auth.parameters.AccessToken)),
            [Microsoft.VisualStudio.Services.Common.CredentialPromptType]::DoNotPrompt)

        # Add-Type -LiteralPath (Assert-Path "$(Get-TaskVariable -Name 'Agent.ServerOMDirectory' -Require)\Microsoft.VisualStudio.Services.Common.dll" -PassThru)
        # $endpoint = (Get-Endpoint -Name SystemVssConnection -Require)
        # New-Object Microsoft.VisualStudio.Services.Common.VssServiceIdentityCredential(
        #     New-Object Microsoft.VisualStudio.Services.Common.VssServiceIdentityToken([string]$endpoint.auth.parameters.AccessToken))
    } finally {
        Trace-LeavingInvocation -InvocationInfo $MyInvocation
    }
}

function New-VssHttpClient {
    [CmdletBinding()]
    param(
        [string]$OMDirectory,

        [Parameter(Mandatory = $true)]
        [string]$TypeName,

        [string]$Uri,

        $VssCredentials)

    Trace-EnteringInvocation -InvocationInfo $MyInvocation
    try {
        if (!$Uri) {
            # Default the URI.
            $Uri = Get-TaskVariable -Name System.TeamFoundationCollectionUri -Require
        }

        [uri]$Uri = New-Object System.Uri($Uri)

        if (!$VssCredentials) {
            # Default the credentials.
            $VssCredentials = Get-VssCredentials
        }

        # Try to load the type.
        try {
            $null = [type]$TypeName
        } catch {
            Write-Verbose "Caught exception while attempting to load type: '$TypeName'"
            if ($TypeName -like 'Microsoft.*.WebApi.*HttpClient') {
                # Try to interpret the DLL name. Trim the last ".___HttpClient" and add ".dll".
                $dll = $TypeName.SubString(0, $TypeName.LastIndexOf('.'))
                $dll = [System.IO.Path]::Combine($OMDirectory, "$dll.dll")
                Write-Verbose "Testing file path for interpreted assembly name: '$dll'"
                if (!(Test-Path -LiteralPath $dll -PathType Leaf)) {
                    Write-Verbose 'Not found. Rethrowing exception.'
                    # Unable to find a file matching the interpreted assembly name.
                    # Rethrow.
                    throw
                }

                # Load the interpreted WebApi DLL.
                Write-Verbose "Loading assembly: '$dll'"
                Add-Type -LiteralPath $dll

                # Try again to load the type, now that the WebApi DLL is loaded.
                Write-Verbose "Re-attempting to load the type: '$TypeName'"
                $null = [type]$TypeName
            } else {
                # Rethrow.
                throw
            }
        }

        # Try to construct the HTTP client.
        try {
            New-Object $TypeName($Uri, $VssCredentials)
        } catch {
            # Check if the exception is due to Newtonsoft.Json DLL not found.
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
                New-Object $TypeName($Uri, $VssCredentials)
            } finally {
                # Unregister the assembly resolver.
                [System.AppDomain]::CurrentDomain.remove_AssemblyResolve($onAssemblyResolve)
            }
            # '*'*80|out-host
            # $_.exception.gettype().fullname|out-host
            # '*'*80|out-host
            # $_.exception|fl * -for|out-host
            # '*'*80|out-host
            # $_.exception.innerexception.gettype().fullname|out-host
            # '*'*80|out-host
            # $_.exception.innerexception|fl * -for|out-host
            # '*'*80|out-host
        }
    } finally {
        Trace-LeavingInvocation -InvocationInfo $MyInvocation
    }
}