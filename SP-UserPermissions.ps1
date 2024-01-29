# Connect to SharePoint Online
Connect-SPOService -Url "https://your-tenant-name-admin.sharepoint.com"

# Specify the user's email address
$userEmail = "user@example.com"

# Create an empty array to store the permissions report
$permissionsReport = @()

# Get all site collections in the tenant
$siteCollections = Get-SPOSite

# Loop through each site collection
foreach ($siteCollection in $siteCollections) {
    Write-Host "Site Collection: $($siteCollection.Url)"

    # Get all sites in the site collection
    $sites = Get-SPOSite -Identity $siteCollection.Url -Limit All

    # Loop through each site
    foreach ($site in $sites) {
        Write-Host "  Site: $($site.Url)"

        # Get all document libraries in the site
        $documentLibraries = Get-SPOList -Web $site.Url -IncludeAllProperties | Where-Object { $_.BaseType -eq "DocumentLibrary" }

        # Loop through each document library
        foreach ($documentLibrary in $documentLibraries) {
            Write-Host "    Document Library: $($documentLibrary.Title)"

            # Get the user's permissions for the document library
            $userPermissions = Get-SPOUser -Site $site.Url -LoginName $userEmail -List $documentLibrary.Title

            # Add the user's permissions to the report
            $permissionsReport += $userPermissions | Select-Object SiteUrl, UserLoginName, RoleAssignment, MemberType

            # Get all files in the document library
            $files = Get-SPOListItem -List $documentLibrary.Title -Web $site.Url

            # Loop through each file
            foreach ($file in $files) {
                Write-Host "      File: $($file.FieldValues.FileLeafRef)"

                # Get the user's permissions for the file
                $filePermissions = Get-SPOUser -Site $site.Url -LoginName $userEmail -List $documentLibrary.Title -Item $file.FieldValues.ID

                # Add the user's permissions to the report
                $permissionsReport += $filePermissions | Select-Object SiteUrl, UserLoginName, RoleAssignment, MemberType
            }
        }
    }
}

# Generate the permissions report
$permissionsReport | Export-Csv -Path "PermissionsReport.csv" -NoTypeInformation
