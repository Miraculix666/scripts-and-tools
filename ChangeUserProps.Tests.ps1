BeforeAll {
    $scriptPath = Join-Path $PSScriptRoot "ChangeUserProps.ps1"

    # Extract the Test-ADModule function using AST
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
    $functionAst = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Test-ADModule'
    }, $true)

    if (-not $functionAst) {
        throw "Function Test-ADModule not found in $scriptPath"
    }

    Invoke-Expression $functionAst[0].Extent.Text
}

Describe "Test-ADModule" {
    Context "When ActiveDirectory module is not available" {
        BeforeAll {
            Mock Get-Module { return $null } -ParameterFilter { $ListAvailable }
            Mock Write-Verbose {}
            Mock Write-Error {}
        }

        It "Returns false and writes an error" {
            $result = Test-ADModule

            $result | Should -Be $false
            Should -Invoke -CommandName Get-Module -Times 1 -ParameterFilter { $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Should -Invoke -CommandName Write-Error -Times 1 -ParameterFilter { $Message -like '*nicht verfügbar*' }
        }
    }

    Context "When module is available but not loaded, and Import-Module fails" {
        BeforeAll {
            Mock Get-Module { return $true } -ParameterFilter { $ListAvailable }
            Mock Get-Module { return $null } -ParameterFilter { -not $ListAvailable }
            Mock Import-Module { throw "Mock Error" }
            Mock Write-Verbose {}
            Mock Write-Error {}
        }

        It "Returns false and writes an error" {
            $result = Test-ADModule

            $result | Should -Be $false
            Should -Invoke -CommandName Get-Module -Times 1 -ParameterFilter { $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Should -Invoke -CommandName Get-Module -Times 1 -ParameterFilter { -not $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Should -Invoke -CommandName Import-Module -Times 1 -ParameterFilter { $Name -eq 'ActiveDirectory' }
            Should -Invoke -CommandName Write-Error -Times 1 -ParameterFilter { $Message -like '*Fehler beim Laden*' }
        }
    }

    Context "When module is available, not loaded initially, and loads successfully" {
        BeforeAll {
            Mock Get-Module { return $true } -ParameterFilter { $ListAvailable }
            Mock Get-Module { return $null } -ParameterFilter { -not $ListAvailable }
            Mock Import-Module {}
            Mock Write-Verbose {}
            Mock Write-Error {}
        }

        It "Returns true and writes verbose success message" {
            $result = Test-ADModule

            $result | Should -Be $true
            Should -Invoke -CommandName Get-Module -Times 1 -ParameterFilter { $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Should -Invoke -CommandName Get-Module -Times 1 -ParameterFilter { -not $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Should -Invoke -CommandName Import-Module -Times 1 -ParameterFilter { $Name -eq 'ActiveDirectory' }
            Should -Invoke -CommandName Write-Verbose -Times 1 -ParameterFilter { $Message -like '*erfolgreich geladen*' }
            Should -Invoke -CommandName Write-Error -Times 0
        }
    }

    Context "When module is already loaded" {
        BeforeAll {
            Mock Get-Module { return $true } -ParameterFilter { $ListAvailable }
            Mock Get-Module { return $true } -ParameterFilter { -not $ListAvailable }
            Mock Import-Module {}
            Mock Write-Verbose {}
            Mock Write-Error {}
        }

        It "Returns true and writes verbose already loaded message" {
            $result = Test-ADModule

            $result | Should -Be $true
            Should -Invoke -CommandName Get-Module -Times 1 -ParameterFilter { $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Should -Invoke -CommandName Get-Module -Times 1 -ParameterFilter { -not $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Should -Invoke -CommandName Import-Module -Times 0
            Should -Invoke -CommandName Write-Verbose -Times 1 -ParameterFilter { $Message -like '*bereits geladen*' }
            Should -Invoke -CommandName Write-Error -Times 0
        }
    }
}
