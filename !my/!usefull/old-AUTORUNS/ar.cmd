@echo off
: autorunsc.exe -a l -h -s -t -nobanner -o \\srv-ad\d$\audit\ar-srv-term.txt

psexec64 -accepteula -nobanner -c -e -f \\srv-term autorunsc.exe -a l -h -s -t -nobanner -o ar-srv-term.txt
move /Y \\srv-term\admin$\system32\ar-srv-term.txt \\srv-ad\d$\audit\ar-srv-term.txt
:del \\srv-term\admin$\system32\ar-srv-term.txt