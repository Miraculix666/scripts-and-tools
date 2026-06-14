BeforeAll {
    # Mocking Import-Module to prevent exit 1 if ActiveDirectory is not installed
    Mock Import-Module {}

    $ScriptPath = Join-Path -Path $PSScriptRoot -ChildPath 'UserCopyCreate.ps1'
    . $ScriptPath
}

Describe 'New-CustomADUser' {
    BeforeEach {
        # General Mocks for Cmdlets used in New-CustomADUser
        Mock Get-ADOrganizationalUnit { return $true }
        Mock New-ADUser {}
        Mock Add-ADGroupMember {}
        Mock Write-CustomLog {}
    }

    It 'Should create a user successfully with standard parameters' {
        $params = @{
            SamAccountName = 'jdoe'
            UserPrincipalName = 'jdoe@domain.local'
            Name = 'John Doe'
            GivenName = 'John'
            Surname = 'Doe'
            OU = 'OU=Users,DC=domain,DC=local'
            Password = (ConvertTo-SecureString 'TestPassword123!' -AsPlainText -Force)
        }

        New-CustomADUser @params

        Assert-MockCalled Get-ADOrganizationalUnit -Times 1 -ParameterFilter { $Filter.ToString() -eq "DistinguishedName -eq OU=Users,DC=domain,DC=local" }
        Assert-MockCalled New-ADUser -Times 1 -ParameterFilter {
            $Name -eq 'John Doe' -and
            $SamAccountName -eq 'jdoe' -and
            $UserPrincipalName -eq 'jdoe@domain.local' -and
            $Path -eq 'OU=Users,DC=domain,DC=local'
        }
        Assert-MockCalled Add-ADGroupMember -Times 0
        Assert-MockCalled Write-CustomLog -Times 1 -ParameterFilter {
            $Message -match "erfolgreich erstellt"
        }
    }

    It 'Should add user to groups if Groups parameter is provided' {
        $params = @{
            SamAccountName = 'jdoe'
            UserPrincipalName = 'jdoe@domain.local'
            Name = 'John Doe'
            OU = 'OU=Users,DC=domain,DC=local'
            Groups = @('Group1', 'Group2')
        }

        New-CustomADUser @params

        Assert-MockCalled New-ADUser -Times 1
        Assert-MockCalled Add-ADGroupMember -Times 1 -ParameterFilter { $Identity -eq 'Group1' -and $Members -eq 'jdoe' }
        Assert-MockCalled Add-ADGroupMember -Times 1 -ParameterFilter { $Identity -eq 'Group2' -and $Members -eq 'jdoe' }
    }

    It 'Should log a warning if group addition fails' {
        Mock Add-ADGroupMember { throw "Group not found" }

        $params = @{
            SamAccountName = 'jdoe'
            UserPrincipalName = 'jdoe@domain.local'
            Name = 'John Doe'
            OU = 'OU=Users,DC=domain,DC=local'
            Groups = @('InvalidGroup')
        }

        New-CustomADUser @params

        Assert-MockCalled Write-CustomLog -Times 1 -ParameterFilter {
            $Level -eq 'WARNUNG' -and
            $Message -match "Fehler beim Hinzufügen von Benutzer 'jdoe' zu Gruppe 'InvalidGroup'"
        }
    }

    It 'Should throw an error and log if OU does not exist' {
        Mock Get-ADOrganizationalUnit { return $null }

        $params = @{
            SamAccountName = 'jdoe'
            UserPrincipalName = 'jdoe@domain.local'
            Name = 'John Doe'
            OU = 'OU=InvalidOU,DC=domain,DC=local'
        }

        { New-CustomADUser @params } | Should -Throw "Die angegebene OU existiert nicht: OU=InvalidOU,DC=domain,DC=local"

        Assert-MockCalled New-ADUser -Times 0
        Assert-MockCalled Write-CustomLog -Times 1 -ParameterFilter {
            $Level -eq 'FEHLER' -and
            $Message -match "Fehler bei der Erstellung des Benutzers 'John Doe'"
        }
    }

    It 'Should throw an error and log if user creation fails' {
        Mock New-ADUser { throw "Access Denied" }

        $params = @{
            SamAccountName = 'jdoe'
            UserPrincipalName = 'jdoe@domain.local'
            Name = 'John Doe'
            OU = 'OU=Users,DC=domain,DC=local'
        }

        { New-CustomADUser @params } | Should -Throw "Access Denied"

        Assert-MockCalled Write-CustomLog -Times 1 -ParameterFilter {
            $Level -eq 'FEHLER' -and
            $Message -match "Fehler bei der Erstellung des Benutzers 'John Doe'"
        }
    }
}
