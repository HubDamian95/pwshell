start-transcript

$singleDirectory = @("example1","example2"
)

Foreach($dir in $singleDirectory){
write-host "$dir"
$shares = net view \\$dir /all | select -Skip 7 | ?{$_ -match 'disk*'} | %{$_ -match '^(.+?)\s+Disk*'|out-null;$matches[1]}
Foreach($share in $shares){
Write-Host "Listing the share for $dir "
Write-Host "Share mentioned $share"
get-acl \\$dir\$share | fl
}
}

end-trascript