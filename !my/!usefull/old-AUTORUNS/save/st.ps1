
#Get-AdGroup -Filter 'SamAccountName -Like "*�����*"' | Get-AdGroupMember -Recursive | select-object -uniq

Get-AdGroup -Filter 'SamAccountName -Like "*admin*" -Or SamAccountName -like "*�����*"' | Get-AdGroupMember -Recursive | select-object -uniq |out-file aa.txt -Encoding oem