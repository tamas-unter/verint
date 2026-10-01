[xml]$shortcuts=cat "$($env:userprofile)\AppData\Roaming\Notepad++\shortcuts.xml"
if($shortcuts.NotepadPlus.Macros.Macro |where name -eq c0){
    # macro is there already
    Write-Host "Shortcuts already contain my macros :)"
} else {
    # need to copy or create


$content=@'
<NotepadPlus>
    <InternalCommands/>
    <Macros>
        <Macro name="c4" Ctrl="yes" Alt="yes" Shift="yes" Key="53">
            <Action type="2" message="0" wParam="43030" lParam="0" sParam=""/>
        </Macro>
        <Macro name="c3" Ctrl="yes" Alt="yes" Shift="yes" Key="52">
            <Action type="2" message="0" wParam="43028" lParam="0" sParam=""/>
        </Macro>
        <Macro name="c2" Ctrl="yes" Alt="yes" Shift="yes" Key="51">
            <Action type="2" message="0" wParam="43026" lParam="0" sParam=""/>
        </Macro>
        <Macro name="c1" Ctrl="yes" Alt="yes" Shift="yes" Key="50">
            <Action type="2" message="0" wParam="43024" lParam="0" sParam=""/>
        </Macro>
        <Macro name="c0" Ctrl="yes" Alt="yes" Shift="yes" Key="49">
            <Action type="2" message="0" wParam="43022" lParam="0" sParam=""/>
        </Macro>
    </Macros>
    <UserDefinedCommands />
    <PluginCommands />
    <ScintillaKeys />
</NotepadPlus>
'@

# avoiding BOM to screw up with np++
#$content | out-file -Encoding "UTF8" "$($env:userprofile)\AppData\Roaming\Notepad++\shortcuts.xml"
$nobom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllLines("$($env:userprofile)\AppData\Roaming\Notepad++\shortcuts.xml", $content, $nobom)


}
# SIG # Begin signature block
# MIIFcAYJKoZIhvcNAQcCoIIFYTCCBV0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUjfvDNu2Rdw0zwb9R7uLEGQiz
# GRWgggMKMIIDBjCCAe6gAwIBAgIQdwWvG+eWnJpNRaSZ/LknezANBgkqhkiG9w0B
# AQsFADAbMRkwFwYDVQQDDBBBVEEgQXV0aGVudGljb2RlMB4XDTI1MDQwNzE0NTU1
# M1oXDTI2MDQwNzE1MTU1M1owGzEZMBcGA1UEAwwQQVRBIEF1dGhlbnRpY29kZTCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAMvoBTdOhBV57TaEk3zaxaRf
# qLL6QzoNBQByg/HBL8BxcEgTgH3dtgGYGKtLevFr3dm9UXDlsCPjlmtpTnhDu2cy
# n7XZwuBP3wRdCHEyGCPQiU4mfH9LBuVxe3z1r4MuVUqck7btOptqTHziG4xAICuy
# 3lrFZ4gPWiYDwQQXnH/N745Bx5aKEjVrSxxPl6HpxZfg065cwRQ6GvpaUc2s/X7D
# mm00euK7tWF8uj/jF5JGCDbQD51LjtHoAD32arR73Y8ypdFE5VhPyxntuo4Z7RDA
# ANzNnOqYYzljvyntL0Mo/xNiKXYqe0w/2og/cjqxkkODzozDmH/eYSr1tIq8hKUC
# AwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0G
# A1UdDgQWBBQZoVvH8OeAs2iC0G80CAhcQAbF/zANBgkqhkiG9w0BAQsFAAOCAQEA
# xW3Bvt+2cYVw0V7wwFH5huCLwCmFTNFYmpIo4E7WAPFPnhESbTPdT7Dzx0oSyUtf
# 3ijZK3l36a1Ifjsan8/SHyPSU4gMBC98SIWjce1b3wP20H1ZkKhNQ+GUJzE1WgPP
# qoaq+y/azgg/FY/yLU2cTYjrobGaI8/E3br5xDzw98pcQ60PlVRm864hOSiZrefo
# XNuzTveP3y2xMf+oPu50efP8LhIPQtN7hAmt2MfPJaUqKzBFVc3+1N6Yxf1yuVah
# UZtG4xTIeGb7JGv9sSOoJ+If+vXC1C3nDmi7O5JBtb1/QlQOfHbQXgoyzrLWQLZI
# hVOWeC8wXnmomKjdhMQXZzGCAdAwggHMAgEBMC8wGzEZMBcGA1UEAwwQQVRBIEF1
# dGhlbnRpY29kZQIQdwWvG+eWnJpNRaSZ/LknezAJBgUrDgMCGgUAoHgwGAYKKwYB
# BAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAc
# BgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUJP2p
# 4LQMhjhe4rgFGv8TyK0Dg7gwDQYJKoZIhvcNAQEBBQAEggEAMk6Nx0R5GIhJdT4U
# noV92BoSwriFziAAzqMyjvNV924Q26O3MOsutt1GDUGDWMu8IUYTeTK+raqgsZiq
# ub1O8QZmbxuq0M3vHF6fooFfXT9RgMNfiTTaZH6G3xJG5rjxRskqCeI86tr40kkR
# d1Bk94rnQUhdzDAnYXm2baQVxWSVa0hWYs0V+mrI6sJBrMJ/W0tOpU+EOI9b3dJw
# W1qgFUPYDee+k0iAxx9xDCWcy3kIqM+G4HqxgY/Rmz3ypgcknLKLhkPOCrfUouIx
# lNyNW5tyzkLCjpON4/+ftIbLzDIOSYIfu10QwIfw+unVVg5DBtfXXv0OT8BAUC3k
# ZWVl6w==
# SIG # End signature block
