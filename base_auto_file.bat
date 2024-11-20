@ECHO OFF
set file=%1
powershell -sta -command "$sc=New-Object System.Collections.Specialized.StringCollection; $sc.Add('%file%'); Add-Type -Assembly 'System.Windows.Forms'; [System.Windows.Forms.Clipboard]::SetFileDropList($sc)"