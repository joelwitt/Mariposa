# Script to get account information and space usage in Exchange Server 2019
# Import the Exchange module

Add-Pssnapin Microsoft.Exchange.Management.Powershell.SnapIn

# Get the list of mailboxes that belong to the domain

$buzones = Get-Mailbox -ResultSize Unlimited

# Initialize variables to store information

$totalSpaceUsed = 0
$accountDetails = @()

foreach ($buzon in $buzones) {
    $mailboxStats = Get-MailboxStatistics -Identity $buzon.Identity
    $spaceUsed = [math]::Round($mailboxStats.TotalItemSize.Value.ToMB(), 2)
    $totalSpaceUsed += $spaceUsed
    $accountDetails += [PSCustomObject]@{
        AccountName = $buzon.DisplayName
        Email = $buzon.PrimarySmtpAddress
        SpaceUsedMB = $spaceUsed
    }

    # Show progress
    Write-Progress -Activity "Processing mailboxes" -Status "Processing $($buzon.DisplayName)" -PercentComplete (($accountDetails.Count / $buzones.Count) * 100)
}

# Calculate the total space used by all accounts in GB
$totalSpaceUsedGB = [math]::Round($totalSpaceUsed / 1024, 2)

# Display information
Write-Output "Number of accounts: $($buzones.Count)"
Write-Output "Space used per account (MB):"
$accountDetails | Format-Table -AutoSize
Write-Output "Total space used by all accounts: $totalSpaceUsedGB GB"

# Save information to a CSV file
$outputPath = "C:\Temp\account_info.csv"
$accountDetails | Export-Csv -Path $outputPath -NoTypeInformation
Add-Content -Path $outputPath -Value "`"Total space used by all accounts`",$totalSpaceUsedGB`"GB"

Write-Output "Information saved to file: $outputPath"
