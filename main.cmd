:: Simple shim launcher, for compatibility with cmd.exe.
:: -ExecutionPolicy Bypass allows us to run the script without
:: requiring the end-user to change their PowerShell execution
:: policy to something more sensible than the default Restricted.

@powershell -ExecutionPolicy Bypass -File main.ps1 %*
