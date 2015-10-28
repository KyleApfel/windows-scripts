Ping 127.0.0.1 -n 10
 
#Custom variables 
$CompanyName = 'Strickland Propane' 
$DomainName = 'spropane.llc' 
$ADModify = $ADUser.whenChanged

$SigSource = “\\propaneserver\SYSVOL\spropane.llc\scripts\Signature” 

# Forced as New
$ForceSignatureNew = '1' #When the signature are forced the signature are enforced as default signature for new messages the next time the script runs. 0 = no force, 1 = force 
$ForceSignatureReplyForward = '0' #When the signature are forced the signature are enforced as default signature for reply/forward messages the next time the script runs. 0 = no force, 1 = force 
 
#Environment variables 
$AppData=(Get-Item env:appdata).value 
$SigPath = '\Microsoft\Signatures' 
$LocalSignaturePath = $AppData+$SigPath 
$RemoteSignaturePathFull = $SigSource+'\'+$CompanyName+'.docx' 
 
#Get Active Directory information for current user 

$UserName = $env:username 
$Filter = “(&(objectCategory=User)(samAccountName=$UserName))” 
$Searcher = New-Object System.DirectoryServices.DirectorySearcher 
$Searcher.Filter = $Filter 
$ADUserPath = $Searcher.FindOne() 
$ADUser = $ADUserPath.GetDirectoryEntry() 
$ADDisplayName = $ADUser.DisplayName 
$ADEmailAddress = $ADUser.mail 
$ADwwwHomePage = $ADUser.wWWHomePage
$ADTitle = $ADUser.title 
$ADTelephoneNumber = $ADUser.telephoneNumber 
$ADMobileNumber = $ADUser.mobile
#$ADoffice = $ADUser.physicalDeliveryOfficeName
$ADAddress = $ADUser.streetAddress
$ADCity = $ADUser.l
$ADState = $ADUser.st
$ADZip = $ADUser.postalCode
$ADCountry = $ADUser.co

#Setting registry information for the current user 
$CompanyRegPath = “HKCU:\Software\”+$CompanyName 
 
if (Test-Path $CompanyRegPath) 
{} 
else 
{New-Item -path “HKCU:\Software” -name $CompanyName} 
 
if (Test-Path $CompanyRegPath’\Outlook Signature Settings’) 
{} 
else 
{New-Item -path $CompanyRegPath -name “Outlook Signature Settings”} 

If (Test-Path $LocalSignaturePath) {}
else
{New-Item -ItemType directory -Path $LocalSignaturePath}


###choose when to update signature, after user update, or after signature update. ### Change accordingly. 
$SigVersion = (gci $RemoteSignaturePathFull).LastWriteTime #When was the last time the signature was written 
#$SigVersion = $ADModify # checks if active directory user object has been updated

$ForcedSignatureNew = (Get-ItemProperty $CompanyRegPath’\Outlook Signature Settings’).ForcedSignatureNew 
$ForcedSignatureReplyForward = (Get-ItemProperty $CompanyRegPath’\Outlook Signature Settings’).ForcedSignatureReplyForward 
$SignatureVersion = (Get-ItemProperty $CompanyRegPath’\Outlook Signature Settings’).SignatureVersion 
Set-ItemProperty $CompanyRegPath’\Outlook Signature Settings’ -name SignatureSourceFiles -Value $SigSource 
$SignatureSourceFiles = (Get-ItemProperty $CompanyRegPath’\Outlook Signature Settings’).SignatureSourceFiles 
 
#Forcing signature for new messages if enabled 
if ($ForcedSignatureNew -eq '1') 
{ 
#Set company signature as default for New messages 
$MSWord = New-Object -com word.application 
$EmailOptions = $MSWord.EmailOptions 
$EmailSignature = $EmailOptions.EmailSignature 
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries 
$EmailSignature.NewMessageSignature=$CompanyName 
#Added This
$EmailSignature.ReplyMessageSignature=$CompanyName 
$MSWord.Quit() 
} 
 
#Forcing signature for reply/forward messages if enabled 
if ($ForcedSignatureReplyForward -eq '1') 
{ 
#Set company signature as default for Reply/Forward messages 
$MSWord = New-Object -com word.application 
$EmailOptions = $MSWord.EmailOptions 
$EmailSignature = $EmailOptions.EmailSignature 
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries 
$EmailSignature.ReplyMessageSignature=$CompanyName 
$MSWord.Quit() 
} 
 
#Copying signature sourcefiles and creating signature if signature-version are different from local version 
#if ($SignatureVersion -ne $SigVersion){} 
#else 
#{ 
#Copy signature templates from domain to local Signature-folder 
Copy-Item “$SignatureSourceFiles\*” $LocalSignaturePath -Recurse -Force 
 
$ReplaceAll = 2 
$FindContinue = 1 
$MatchCase = $False 
$MatchWholeWord = $True 
$MatchWildcards = $False 
$MatchSoundsLike = $False 
$MatchAllWordForms = $False 
$Forward = $True 
$Wrap = $FindContinue 
$Format = $False 
 
#Insert variables from Active Directory to rtf signature-file 
$MSWord = New-Object -com word.application 
$fullPath = $LocalSignaturePath+'\'+$CompanyName+'.docx' 
$MSWord.Documents.Open($fullPath) 

$FindText = “DisplayName” 
$ReplaceText = $ADDisplayName.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 
 
$FindText = “Title” 
$ReplaceText = $ADTitle.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 

$FindText = “MobileNumber” 
$ReplaceText = $ADMobileNumber.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 

$FindText = “TelephoneNumber” 
$ReplaceText = $ADTelephoneNumber.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 

#$FindText = “office” 
#$ReplaceText = $ADoffice.ToString() 
#$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 

$FindText = “Address” 
$ReplaceText = $ADAddress.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 

$FindText = “City” 
$ReplaceText = $ADCity.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 

$FindText = “State” 
$ReplaceText = $ADState.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 
 
$FindText = “Zip” 
$ReplaceText = $ADZip.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 
 
$FindText = “Country” 
$ReplaceText = $ADCountry.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 

$MSWord.Selection.Find.Execute(“EmailAddress”) 
$MSWord.ActiveDocument.Hyperlinks.Add($MSWord.Selection.Range, “mailto:”+$ADEmailAddress.ToString(), $missing, $missing, $ADEmailAddress.ToString())  

$MSWord.Selection.Find.Execute(“wwwHomePage”) 
$MSWord.ActiveDocument.Hyperlinks.Add($MSWord.Selection.Range, ""+$ADwwwHomePage.ToString(), $missing, $missing, $ADwwwHomePage.ToString())  
 
$MSWord.ActiveDocument.Save() 

#$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], “wdFormatHTML”); 
[ref]$BrowserLevel = “microsoft.office.interop.word.WdBrowserLevel” -as [type] 

$MSWord.ActiveDocument.WebOptions.OrganizeInFolder = $true 
$MSWord.ActiveDocument.WebOptions.UseLongFileNames = $true 
$MSWord.ActiveDocument.WebOptions.BrowserLevel = $BrowserLevel::wdBrowserLevelMicrosoftInternetExplorer6 

#Fixes Enumeration Problems
$wdTypes = Add-Type -AssemblyName 'Microsoft.Office.Interop.Word' -Passthru
$wdSaveFormat = $wdTypes | Where {$_.Name -eq "wdSaveFormat"}
  
#Save HTML
$path = $LocalSignaturePath+'\'+$CompanyName+".htm"
$MSWord.ActiveDocument.saveas([ref]$path, [ref]$wdSaveFormat::wdFormatHTML);
$MSWord.ActiveDocument.saveas($path, $wdSaveFormat::wdFormatHTML);
    
#Save RTF 
$path = $LocalSignaturePath+'\'+$CompanyName+".rtf"
$MSWord.ActiveDocument.SaveAs([ref]$path, [ref]$wdSaveFormat::wdFormatRTF);
$MSWord.ActiveDocument.SaveAs($path, $wdSaveFormat::wdFormatRTF);
  
#Save TXT    
$path = $LocalSignaturePath+'\'+$CompanyName+".txt"
$MSWord.ActiveDocument.SaveAs([ref]$path, [ref]$wdSaveFormat::wdFormatText);
$MSWord.ActiveDocument.SaveAs($path, $wdSaveFormat::wdFormatText);
  
$MSWord.ActiveDocument.Close();
$MSWord.Quit();
 
#} 
 
#Stamp registry-values for Outlook Signature Settings if they doesn`t match the initial script variables. Note that these will apply after the second script run when changes are made in the “Custom variables”-section. 
if ($ForcedSignatureNew -eq $ForceSignatureNew){} 
else 
{Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name ForcedSignatureNew -Value $ForceSignatureNew} 
 
if ($ForcedSignatureReplyForward -eq $ForceSignatureReplyForward){} 
else 
{Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name ForcedSignatureReplyForward -Value $ForceSignatureReplyForward} 
 
if ($SignatureVersion -eq $SigVersion){} 
else 
{Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name SignatureVersion -Value $SigVersion}

