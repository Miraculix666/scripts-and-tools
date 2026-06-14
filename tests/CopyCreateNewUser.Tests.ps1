$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Parent $here) + "\CopyCreateNewUser.ps1"

# To avoid running the entire script on dot-source, we will mock the main logic or just dot source
# but usually it's better to isolate function testing. Since the script executes immediately,
# we'll mock Write-LogMessage, Read-Host, Import-Csv, and Get-ADOrganizationalUnit to prevent real execution.

Describe "Export-TemplateUserData" {
    BeforeAll {
        # Dot source the script while mocking execution
        Mock Write-LogMessage {}
        Mock Read-Host {}
        Mock Import-Csv {}
        Mock Get-ADOrganizationalUnit { return $true }

        # Load functions
        . $sut
    }

    Context "When exporting a valid user" {
        It "Should retrieve the user, select properties, and export to CSV" {
            # Arrange
            $dummyUser = [PSCustomObject]@{
                SamAccountName = "testuser"
                GivenName = "Test"
                Surname = "User"
                Department = "IT"
            }
            $targetPath = ".\test_export.csv"

            Mock Get-ADUser { return $dummyUser }
            Mock Export-Csv {}
            Mock Write-LogMessage {}

            # Act
            Export-TemplateUserData -TemplateUser "testuser" -CsvPath $targetPath

            # Assert
            Assert-MockCalled Get-ADUser -Times 1 -ParameterFilter { $Identity -eq "testuser" }
            Assert-MockCalled Export-Csv -Times 1 -ParameterFilter { $Path -eq $targetPath -and $Delimiter -eq ";" }
            Assert-MockCalled Write-LogMessage -Times 1 -ParameterFilter { $Type -eq "Success" }
        }

        It "Should use the default CsvPath if none is provided" {
            # Arrange
            $dummyUser = [PSCustomObject]@{ SamAccountName = "testuser" }
            Mock Get-ADUser { return $dummyUser }
            Mock Export-Csv {}
            Mock Write-LogMessage {}

            # Act
            Export-TemplateUserData -TemplateUser "testuser"

            # Assert
            Assert-MockCalled Export-Csv -Times 1 -ParameterFilter { $Path -eq ".\TemplateUser_Export.csv" }
        }
    }

    Context "When Get-ADUser fails" {
        It "Should log an error and rethrow" {
            # Arrange
            Mock Get-ADUser { throw "User not found" }
            Mock Write-LogMessage {}

            # Act & Assert
            { Export-TemplateUserData -TemplateUser "invaliduser" } | Should -Throw

            Assert-MockCalled Write-LogMessage -Times 1 -ParameterFilter { $Type -eq "Error" }
        }
    }
}
