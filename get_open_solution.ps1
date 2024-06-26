# get_open_solution.ps1

# Attempt to find the open solution in Visual Studio
$visualStudio = Get-Process devenv -ErrorAction SilentlyContinue

if ($null -ne $visualStudio) {
    try {
        $dte = [Runtime.InteropServices.Marshal]::GetActiveObject("VisualStudio.DTE.17.0")
        if ($dte.Solution.IsOpen) {
            Write-Output $dte.Solution.FullName
        } else {
            Write-Output "No open solution found."
        }
    } catch {
        Write-Output "Error accessing Visual Studio DTE: $_"
    }
} else {
    Write-Output "No Visual Studio instance found."
}
