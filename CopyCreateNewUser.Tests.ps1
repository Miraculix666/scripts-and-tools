# Requires -Module Pester

$ScriptDir = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $ScriptDir "CopyCreateNewUser.ps1")

Describe "New-ADUserFromTemplate" {
    Context "When creating a new user successfully" {
        It "Calls New-ADUser with correct parameters and logs success" {
            # Arrange
            $MockUserData = @{
                SamAccountName = "jdoe"
                GivenName = "John"
                Surname = "Doe"
                Password = "SuperSecretPassword123!"
            }
            $MockTemplateUser = "template.user"
            $MockTargetOU = "OU=Test,DC=Domain,DC=Com"
            $MockDomainObj = [PSCustomObject]@{ DNSRoot = "domain.com" }
            $MockTemplateObj = [PSCustomObject]@{ SamAccountName = "template.user" }

            Mock Get-ADUser { return $MockTemplateObj }
            Mock Get-ADDomain { return $MockDomainObj }
            Mock ConvertTo-SecureString { return "SECURESTRING" }
            Mock New-ADUser { }
            Mock Write-LogMessage { }

            # Act
            New-ADUserFromTemplate -UserData $MockUserData -TemplateUser $MockTemplateUser -TargetOU $MockTargetOU

            # Assert
            Assert-MockCalled New-ADUser -Times 1 -ParameterFilter {
                $SamAccountName -eq "jdoe" -and
                $UserPrincipalName -eq "jdoe@domain.com" -and
                $Name -eq "John Doe" -and
                $Path -eq $MockTargetOU
            }
            Assert-MockCalled Write-LogMessage -Times 1 -ParameterFilter {
                $Message -match "erfolgreich erstellt"
            }
        }
    }

    Context "When New-ADUser fails" {
        It "Catches the exception and logs an error" {
            # Arrange
            $MockUserData = @{
                SamAccountName = "jdoe"
                GivenName = "John"
                Surname = "Doe"
                Password = "SuperSecretPassword123!"
            }
            $MockTemplateUser = "template.user"
            $MockTargetOU = "OU=Test,DC=Domain,DC=Com"

            Mock Get-ADUser { return [PSCustomObject]@{ SamAccountName = "template.user" } }
            Mock Get-ADDomain { return [PSCustomObject]@{ DNSRoot = "domain.com" } }
            Mock ConvertTo-SecureString { return "SECURESTRING" }
            Mock New-ADUser { throw "AD Error" }
            Mock Write-LogMessage { }

            # Act
            New-ADUserFromTemplate -UserData $MockUserData -TemplateUser $MockTemplateUser -TargetOU $MockTargetOU

            # Assert
            Assert-MockCalled New-ADUser -Times 1
            Assert-MockCalled Write-LogMessage -Times 1 -ParameterFilter {
                $Type -eq "Error" -and $Message -match "Fehler beim Erstellen von"
            }
        }
    }
}
