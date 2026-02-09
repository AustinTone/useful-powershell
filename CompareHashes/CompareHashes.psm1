$script:DefaultAlgorithm = "SHA256"

function Compare-HashString {
    param(
        [string] $File,
        [string] $HashString,
        [string] $Algorithm = $script:DefaultAlgorithm
    )

    if (!(test-path $File)){
        Write-Error "File not found!"
    }
    else {
        $filehash = Get-FileHash -Path $File -Algorithm $Algorithm
        if ($filehash.hash -eq $HashString){
            Write-Host "Hashes match" -ForegroundColor Green
        }
        else {
            Write-Warning "Hashes DO NOT match! Proceed with caution."
        }
    }

}

function Compare-TwoFiles {
    param(
        [string] $Original,
        [string] $Diff,
        [string] $Algorithm = $script:DefaultAlgorithm
    )

    if (!(Test-Path $Original) -or !(Test-Path $Diff)){
        Write-Error "One or both files cannot be found!"
    }
    else {
        $origHash = Get-FileHash -Path $Original -Algorithm $Algorithm
        $diffHash = Get-FileHash -Path $Diff -Algorithm $Algorithm

        if ($diffHash.Hash -eq $origHash.Hash){
            Write-Host "Both hashes match!" -ForegroundColor Green
        }
        else {
            Write-Warning "Hashes DO NOT match! Proceed with caution."
        }
    }
}

Export-ModuleMember -Function Compare-HashString, Compare-TwoFiles