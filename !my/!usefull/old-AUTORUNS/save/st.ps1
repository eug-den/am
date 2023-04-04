
#Get-AdGroup -Filter 'SamAccountName -Like "*Админ*"' | Get-AdGroupMember -Recursive | select-object -uniq

Get-AdGroup -Filter 'SamAccountName -Like "*admin*" -Or SamAccountName -like "*Админ*"' | Get-AdGroupMember -Recursive | select-object -uniq |out-file aa.txt -Encoding oem