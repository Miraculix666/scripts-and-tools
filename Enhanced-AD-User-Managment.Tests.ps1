$ErrorActionPreference = "Stop"

# Pester 5 syntax
Describe "Apply-ADUserProperties" {
    BeforeAll {
        # Extract the function 'Apply-ADUserProperties' from the script into memory.
        $scriptPath = Join-Path $PSScriptRoot "Enhanced-AD-User-Managment.ps1"
        $scriptContent = Get-Content $scriptPath -Raw

        # In PowerShell, we cannot use [ref]$null, we must provide initialized variables
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($scriptContent, [ref]$tokens, [ref]$errors)

        $functionAst = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq 'Apply-ADUserProperties' }, $true)

        if ($functionAst.Count -gt 0) {
            Invoke-Expression $functionAst[0].Extent.Text
        } else {
            throw "Function Apply-ADUserProperties not found in script."
        }
    }

    Context "When reference user has new/updated properties" {
        BeforeEach {
            # Mock ADUser object creation
            $referenceUser = New-Object PSObject -Property @{
                SamAccountName = "RefUser"
                Description = "New Description"
                Office = "Office 1"
                Department = $null
            }
            $referenceUser.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADUser')

            $targetUser = New-Object PSObject -Property @{
                SamAccountName = "TargetUser"
                Description = "Old Description"
                Office = "Office 1"
                Department = "HR"
            }
            $targetUser.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADUser')

            # Mocks
            Mock Write-Log {}
            Mock Write-Verbose {}
            Mock Add-UserReportEntry {}
            Mock Set-ADUser {}

            # Groups mock
            Mock Get-ADPrincipalGroupMembership {
                if ($Identity.SamAccountName -eq "RefUser") {
                    return @(
                        New-Object PSObject -Property @{ Name = "Group A"; DistinguishedName = "CN=Group A" }
                    )
                } else {
                    return @()
                }
            }

            Mock Compare-Object {
                return @(
                    New-Object PSObject -Property @{ Name = "Group A"; DistinguishedName = "CN=Group A"; SideIndicator = "<=" }
                )
            }

            Mock Add-ADPrincipalGroupMembership {}
        }

        It "Should call Set-ADUser with changed properties only (excluding nulls)" {
            # Arrange & Act
            Apply-ADUserProperties -ReferenceUser $referenceUser -TargetUser $targetUser

            # Assert using Pester 5 syntax
            Should -Invoke -CommandName Set-ADUser -Times 1 -ParameterFilter {
                $Identity.SamAccountName -eq 'TargetUser' -and
                $Description -eq 'New Description'
            }

            # Verify Office wasn't included in the change because it was identical
            # And Department wasn't included because it's null on the reference user
        }

        It "Should add missing group memberships from reference user" {
            # Arrange & Act
            Apply-ADUserProperties -ReferenceUser $referenceUser -TargetUser $targetUser

            # Assert using Pester 5 syntax
            Should -Invoke -CommandName Add-ADPrincipalGroupMembership -Times 1 -ParameterFilter {
                $Identity.SamAccountName -eq 'TargetUser'
            }
        }
    }

    Context "When reference user has identical properties and no new groups" {
        BeforeEach {
            $identicalUser = New-Object PSObject -Property @{
                SamAccountName = "SameUser"
                Description = "Same"
            }
            $identicalUser.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADUser')

            Mock Write-Log {}
            Mock Write-Verbose {}
            Mock Add-UserReportEntry {}
            Mock Set-ADUser {}

            Mock Get-ADPrincipalGroupMembership { return @() }
            Mock Compare-Object { return @() }
            Mock Add-ADPrincipalGroupMembership {}
        }

        It "Should not call Set-ADUser" {
            Apply-ADUserProperties -ReferenceUser $identicalUser -TargetUser $identicalUser
            Should -Invoke -CommandName Set-ADUser -Times 0
        }

        It "Should not call Add-ADPrincipalGroupMembership" {
            Apply-ADUserProperties -ReferenceUser $identicalUser -TargetUser $identicalUser
            Should -Invoke -CommandName Add-ADPrincipalGroupMembership -Times 0
        }
    }
}
