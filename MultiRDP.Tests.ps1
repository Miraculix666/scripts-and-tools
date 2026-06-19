BeforeAll {
    $errors = $null
    $tokens = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile("$(Get-Location)/MultiRDP.ps1", [ref]$tokens, [ref]$errors)
    $functionAst = $ast.Find({$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq 'Test-RDPConnection'}, $true)
    if ($functionAst) {
        Invoke-Expression $functionAst.Extent.Text
    } else {
        throw "Function Test-RDPConnection not found in MultiRDP.ps1"
    }
}

Describe "Test-RDPConnection" {
    BeforeEach {
        Mock Write-Progress -MockWith {}
        Mock Write-Host -MockWith {}
        Mock Write-Verbose -MockWith {}
        Mock Format-Table -MockWith {} -ParameterFilter { $true }
        Mock Start-Sleep -MockWith {} -ParameterFilter { $true }
    }

    It "Sollte bei einem leeren Array an Computern einen Fehler werfen" {
        { Test-RDPConnection -Computers @() } | Should -Throw
    }

    It "Sollte das Ergebnis vom Job korrekt durchreichen (1 Computer)" {
        Mock Start-Job -MockWith {
            return [PSCustomObject]@{ Id = 1; State = 'Completed' }
        }
        Mock Receive-Job -MockWith {
            return [PSCustomObject]@{
                ComputerName = "server01"
                DNSStatus = "Aufgelöst"
                PingStatus = "Erreichbar"
                RDPPortStatus = "Geöffnet"
            }
        } -ParameterFilter { $true }
        Mock Remove-Job -MockWith {} -ParameterFilter { $true }

        $result = @(Test-RDPConnection -Computers @("server01"))

        $result.Count | Should -Be 1
        $result[0].ComputerName | Should -Be "server01"
        $result[0].DNSStatus | Should -Be "Aufgelöst"
        $result[0].PingStatus | Should -Be "Erreichbar"
        $result[0].RDPPortStatus | Should -Be "Geöffnet"
    }

    It "Sollte mit gemischten Ergebnissen korrekt umgehen (2 Computer)" {
        Mock Start-Job -MockWith {
            return [PSCustomObject]@{ Id = 1; State = 'Completed' }
        }
        $global:receiveCallCount = 0
        Mock Receive-Job -MockWith {
            $global:receiveCallCount++
            # Since pipeline sends one item per Job object, and we simulated 2 Jobs by calling Start-Job twice,
            # Receive-Job might be called twice or once with an array. In the mock, we can just track it.
            if ($global:receiveCallCount -eq 1) {
                return [PSCustomObject]@{
                    ComputerName = "server01"
                    DNSStatus = "Aufgelöst"
                    PingStatus = "Erreichbar"
                    RDPPortStatus = "Geöffnet"
                }
            } else {
                return [PSCustomObject]@{
                    ComputerName = "server02"
                    DNSStatus = "Nicht aufgelöst"
                    PingStatus = "Nicht erreichbar"
                    RDPPortStatus = "Geschlossen"
                }
            }
        } -ParameterFilter { $true }
        Mock Remove-Job -MockWith {} -ParameterFilter { $true }

        $result = @(Test-RDPConnection -Computers @("server01", "server02"))

        $result.Count | Should -Be 2

        $server01 = $result | Where-Object { $_.ComputerName -eq "server01" }
        $server01.DNSStatus | Should -Be "Aufgelöst"
        $server01.PingStatus | Should -Be "Erreichbar"
        $server01.RDPPortStatus | Should -Be "Geöffnet"

        $server02 = $result | Where-Object { $_.ComputerName -eq "server02" }
        $server02.DNSStatus | Should -Be "Nicht aufgelöst"
        $server02.PingStatus | Should -Be "Nicht erreichbar"
        $server02.RDPPortStatus | Should -Be "Geschlossen"
    }
}
