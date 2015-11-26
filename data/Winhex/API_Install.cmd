@prompt -$G

REG ADD "HKCU\Software\X-Ways AG\WinHex" /v Path /t REG_SZ /d %CD% /f

@pause