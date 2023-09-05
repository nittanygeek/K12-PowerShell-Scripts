# IMPORTANT: Open and mount the PST in Outlook before running the script
# Generate a CSV file from a mounted PST file in Outlook containing chat messages
# Usage: ./Generate-ChatLogCSV.ps1 -PSTPath C:\path\to\target.pst -CSVPath C:\path\to\destination\file.csv

param(
    [Parameter(Mandatory=$true)]
    [string]$PSTPath,
    [Parameter(Mandatory=$true)]
    [string]$CSVPath
)

# Load the Outlook COM object
$Outlook = New-Object -ComObject Outlook.Application

# Load the PST file
$PST = $Outlook.Session.Stores | Where-Object { $_.FilePath -eq $PSTPath }

# Load the root folder of the PST file
$RootFolder = $PST.GetRootFolder()

# Load the folder containing the chat messages.  They are in a folder named "username@domain.com (Primary)" under a folder named "TeamsMessagesData"
$ChatFolder = $RootFolder.Folders | Where-Object { $_.Name -eq "$($PST.DisplayName) (Primary)" } | ForEach-Object { $_.Folders | Where-Object { $_.Name -eq "TeamsMessagesData" } }

# Load the chat messages
$ChatMessages = $ChatFolder.Items | Where-Object { $_.MessageClass -eq "IPM.SkypeTeams.Message" }

# Create a new CSV file
$CSV = New-Object -TypeName System.IO.StreamWriter -ArgumentList $CSVPath

# Write the header row
$CSV.WriteLine("Date,From,To,Message")

# Loop through the chat messages
foreach ($ChatMessage in $ChatMessages) {

    # Extract only the message between the body tags
    $HTML = New-Object -Com "HTMLFile"
    [string]$htmlBody = $ChatMessage.HTMLBody

    try {
        # This works in PowerShell with Office installed
        $html.IHTMLDocument2_write($htmlBody)
    }
    catch {
        # This works when Office is not installed    
        $src = [System.Text.Encoding]::Unicode.GetBytes($htmlBody)
        $html.write($src)
    }

    $MessageBody = $HTML.body.innerText

    if ($MessageBody -eq $null) {
        $MessageBody = ""
    }

    # Write the message to the CSV file
    $CSV.WriteLine("$($ChatMessage.ReceivedTime),$($ChatMessage.SenderName),$($ChatMessage.To),$($MessageBody)")
}

# Close the CSV file
$CSV.Close()

# Close Outlook
$Outlook.Quit()
