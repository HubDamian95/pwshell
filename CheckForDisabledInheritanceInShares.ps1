#Creates a colletion of strings and is then available to a foreach loop at the main function of the program
$shares = @("\\share1", "\\share2")

Write-Output "`nNoninheritable permissions:`n" 

#Logs output to a certain text file
function logOutput($logItem) {

    $logItem | Add-Content outputaaonly.log

}

#Logs console to a certain text file
function logConsole($logItem) {

    Write-Host $logItem -Fore GREEN

}

#Logs console Errors to a log file
function logConsoleError($logItem) {

    Write-Host $logItem -Fore RED

}

#Starts the total transcript
start-transcript -path transcriptapplied.log

#Logs the console log
logConsole("Starting Transcript")


##########################################
#
# Navigates to a required directory and recurses through each one
#
##########################################


foreach ($share in $shares){
dir "" -Directory -Recurse | ForEach-Object {
    $Path = $_.FullName
    try {
logConsole("Processing $($path)")
logOutput("Processing $($path)")
        $TotalACLs = (Get-Acl $Path | select -ExpandProperty Access).Count
        $InheritedCount = (Get-Acl $Path | select -ExpandProperty Access | where { $_.IsInherited -eq $false } | Add-Member -MemberType NoteProperty -Name Path -Value $Path -PassThru | Select Path).Count
        if ($InheritedCount) {
            Write-Output $InheritedCount" out of "$TotalACLs" in "$Path
        }
    }
    catch {
        Write-Error $_
        logOutput("Permission in $Path is not inherited")
        logConsoleError("Permission in $Path is not inherited")
        }
    }
}

Stop-Transcript
logConsole("Transcript Stopped")