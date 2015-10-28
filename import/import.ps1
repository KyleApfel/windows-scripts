# This is a very simple import script using the name column to refer to the CN or Display Name
# for changing whatever column you desire. Look into Set-ADUser to see what fields you can update
# in AD. Update $USERS to choose the location of your CSV. If you're pulling this from my github
# https://github.com/KyleApfel there should be an example CSV inside.

Import-Module ActiveDirectory

$USERS = Import-CSV C:\scripts\csv\googleexport.csv

$USERS|Foreach{
$NAME = $_.name

Get-ADUser -filter { cn -eq $NAME }|Set-ADUser -MobilePhone $_.mobile
Get-ADUser -filter { cn -eq $NAME }|Set-ADUser -Title $_.title

}

