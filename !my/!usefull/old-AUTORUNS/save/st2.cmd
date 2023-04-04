@echo off

set ndate=%date:~-4%%date:~3,2%%date:~0,2%
: set ndate=%date:~-4%%date:~3,2%%date:~0,2%%time:~3,2%%time:~6,2%

@powershell -file st.ps1

rem  }
rem | Export-csv -path out.csv -Encoding OEM}