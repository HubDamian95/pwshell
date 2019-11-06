$array_with_emails = @(
"example@domain.com",
"example2@domain.com")


Foreach($email in $array_with_emails){
write-host "adding $email"
add-distributionGroupMember -identity "exampleGroup" -Member $email -confirm: $false
}
