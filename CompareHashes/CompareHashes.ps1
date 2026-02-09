param(
    #--- Compare to a hash string ---
    # This parameter tells PS to run the string hash comparison function
    [Parameter(Mandatory, ParameterSetName = 'ByString')]
    [switch] $CompareString,

    [Parameter(Mandatory, ParameterSetName = 'ByString')]
    [string] $File,

    [Parameter(Mandatory, ParameterSetName = 'ByString')]
    [string]$HashString,


    #--- Compare to another file ---
    # Tells PS to use the file comparison function
    [Parameter(Mandatory, ParameterSetName = 'ByFile')]
    [switch] $CompareFiles,

    [Parameter(Mandatory, ParameterSetName = "ByFile")]
    [string] $Original,

    [Parameter(Mandatory, ParameterSetName = "ByFile")]
    [string] $Diff,

    # Allows the script to process different hash algorithms but defaults to SHA256
    [string] $Algorithm = "SHA256"
)

begin {
    function Invoke-StringCompare {
        if (!(Test-Path $File)){
            Write-Error "File not found!"
            exit
        }
        else {
            $filehash = Get-FileHash -Path $File -Algorithm $Algorithm

            if ($filehash.Hash -eq $HashString){
                Write-Host "Hashes match!" -ForegroundColor Green
            }
            else {
                Write-Warning "Hashes DO NOT match!!! Proceed with caution!"
            }
        }
    }

    function Invoke-FileCompare {
        if (!(Test-Path -Path $Original) -or !(Test-Path -Path $Diff)){
            Write-Error "One or both files cannot be found!"
        }
        else {
            $origHash = (Get-FileHash -Path $Original -Algorithm $Algorithm).Hash
            $diffHash = (Get-FileHash -Path $Diff -Algorithm $Algorithm).Hash

            if ($origHash -eq $diffHash){
            Write-Host "Both hashes match!"
            }
            else {
                Write-Warning "Hashes DO NOT match!!! Proceed with caution!"
            }
        }
    }
}

process{
    switch ($PSCmdlet.ParameterSetName){
        'ByString' { Invoke-StringCompare }
        'ByFile' { Invoke-FileCompare }
    }
}