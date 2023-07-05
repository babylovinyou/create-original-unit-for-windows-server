# Install the ImportExcel module if not already installed
Install-Module -Name ImportExcel -Scope CurrentUser

# Import the module
Import-Module -Name ImportExcel

# Set the path to the Excel file containing OU information
$excelFilePath = "C:\it\ougroupuser.xlsx"

# Set the worksheet name within the Excel file
$worksheetName = "Sheet1"

# Load the OU data from the Excel file
$ouData = Import-Excel -Path $excelFilePath -WorksheetName $worksheetName

# Import the Active Directory module
Import-Module -Name ActiveDirectory

# Loop through each row in the OU data
foreach ($row in $ouData) {
    $ouName = $row.OUName
    $ouPath = $row.OUPatch

    # Create a new OU
    New-ADOrganizationalUnit -Name $ouName -Path $ouPath -ErrorAction SilentlyContinue
}