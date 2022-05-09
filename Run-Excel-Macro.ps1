function Run-Excel-Macro {
    [CmdletBinding()]

    param (
        [parameter(Mandatory)]
        [ValidateScript( {
                Try {
                    If ((Test-Path -Path $_) -and ($_ -match "\.xla$")) { $True }
                    Else { Throw "$($_) is not a valid Macro path!" }
                }
                Catch {
                    Throw $_
                }
            })]
        [string]$MacroPath,
        [parameter(Mandatory)]
        [ValidateScript( {
                Try {
                    If ((Test-Path -Path $_) -and ($_ -match "\.(xls[xm]?|xml|csv)$")) { $True }
                    Else { Throw "$($_) is not a valid Workbook path!" }
                }
                Catch {
                    Throw $_
                }
            })]
        [string]$workbookPath,
        [parameter(Mandatory)][ValidateNotNullorEmpty()][string]$macroName
    )

        <# 
	.SYNOPSIS

	Loads a macro file and runs a specified macro on the specied Excel

	.DESCRIPTION

    Loads a macro file and runs a specified macro on the specied Excel
    
    .PARAMETER MacroPath
    Specifies the path to the Macro Excel file (*.xla)
    .PARAMETER workbookPath
    Specifies the path to the Macro Excel file (e.g. *.xlsx)
    .PARAMETER macroName
    Specifies the name of the macro

	.EXAMPLE
	PS> . .\Run-Excel-Macro.ps1
	PS> Search-Excel-Report '.\my-macro.xla' '.\my-excel.xlsx' 'macro_name'
#>

    try {
        $MacroPath = Resolve-Path $MacroPath

        $workbookPath = Resolve-Path $workbookPath
        
        $excel = New-Object -comobject Excel.Application

        $macro = $excel.Workbooks.Open($MacroPath)
        $wb = $excel.Workbooks.Open($workbookPath)

        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $excel.activeworkbook.saved = $true
        $excel.Run($macroName)
        $excel.activeworkbook.save()
        $excel.Workbooks.close() 
        $excel.Quit()  

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($macro) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers() 
        Remove-Variable Excel

    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host $_.ScriptStackTrace
    }
}
