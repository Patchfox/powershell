#COPYRIGHT@2020 Mario Schützle
#patchfox.de
#Kontakt: m.schuetzle@patchfox.de


#Alle mit Hash markierten Einträge im Script mit gewünschten Werten füllen 

############################################  MODULE  #####################################################################

#Lädt SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#Modules Installation - braucht man nur einmal : 
#Install-Module -Name Microsoft.Online.SharePoint.PowerShell -RequiredVersion 16.0.8029.0
#Install-Module -Name CredentialManager 
  
Import-Module  Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Import-Module  CredentialManager 


#############################################  ZUGANG  ####################################################################

##Versionen die behalten werden sollen
$VersionsToKeep=# z.B. 3

#Passwort Datei -- Pfad festlegen!
$passpath = # z.B. 'C:\scripts\cred.txt'

#Username eintragen | Admin
$Username = #'UPN' 


#Tenant von Sharepoint eintragen ohne https://
$TenantURL= #'COMPANYNAME.sharepoint.com'

#Schwellenwert Storage zum Abarbeiten in MB
$storagethreshold =#z.B. '10'



######################################  FUNKTION 1  #######################################################

#falls Passwort-Datei noch nicht vorhanden ist
if(!(Test-Path $passpath -PathType Leaf)){
Read-Host -Prompt “Passwort eingeben” -AsSecureString | ConvertFrom-SecureString | Out-File # z.B. “C:\scripts\cred.txt“
}

 

$Pass = Get-Content “C:\scripts\cred.txt” | ConvertTo-SecureString
$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $Username, $Pass


if ($TenantURL -ne '') {

#Manipulation Tenant Admin URL von Tenant URL
$url = "https://"+$TenantURL.Insert($TenantURL.IndexOf("."),"-admin")

#Try{

Connect-SPOService -Url $url -Credential $cred

#Filter von Seiten mit -Filter -like example -notlike example angehängt
$sites = Get-SPOSite -Limit ALL
$results = @()

#Durchläuft alle Seiten und berechnet den genutzen Storage
foreach ($site in $sites) {
$siteStorage = New-Object PSObject
    if(!$sites.StorageUsageCurrent -eq 0 -and !$site.StorageQuota -eq 0) {
    $percent = $site.StorageUsageCurrent / $site.StorageQuota * 100
    }
    else {
    $percent = 0
    }

    $percentage = [math]::Round($percent,2)
    
      $siteStorage | Add-Member -MemberType NoteProperty -Name "Site Title" -Value $site.Title
      $siteStorage | Add-Member -MemberType NoteProperty -Name "Site Url" -Value $site.Url
      $siteStorage | Add-Member -MemberType NoteProperty -Name "Percentage Used" -Value $percentage
      $siteStorage | Add-Member -MemberType NoteProperty -Name "Storage Used (MB)" -Value $site.StorageUsageCurrent
      $siteStorage | Add-Member -MemberType NoteProperty -Name "Storage Quota (MB)" -Value $site.StorageQuota
    
    if($site.StorageUsageCurrent -gt $storagethreshold ) {
    $results += $siteStorage 
    
    }
$siteStorage = $null
}


#################Funktion 2 ###############################################################

##Seiten Position
$position = $null

$systemlibs =@("Converted Forms", "Customized Reports", "Form Templates", "Bilder", "Workflowaufgaben" , 
"Bilder der Websitesammlung", "Formatbibliothek",   "List Template Gallery", "Master Page Gallery", "Pages", 
"Reporting Templates", "Site Assets", "Aktuelle Liste der Site Websiteobjekte",  "Site Collection Documents", "Site Collection Images", "Site Pages", 
"Solution Gallery", "Style Library", "Theme Gallery", "Web Part Gallery", "wfpub", "Inhalts- und Strukturberichte", 
"Formularvorlagen", "Websiteobjekte",  "Registerkarten in Suchergebnissen", "Wiederverwendbarer Inhalt", "Registerkarten auf Suchseiten" ) 


## Anzeigefilter der Seite in XML wie sie nachher abgerufen wird
$qCommand = @"
   <View Scope="Recursive">
        <Where>                         
           <Eq>                  
              <FieldRef Name='FSObjType'/>                                     
              <Value Type='Integer'>0</Value>
            </Eq>               
         </Where>   
              <RowLimit Paged="TRUE">500</RowLimit>
        </View>
"@

$listArray = New-Object System.Collections.Generic.List[System.Object]

#Try {

#Neues Objekt der Credentials MS Sharepoint Client SDK relevant
$Cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $Pass)

$b =0
#Durchlaufe alle Elemente von der vorausgegangenen Schleife aller relevanten Sharepoint Sites
foreach ($entry in $results) {
    #Context
    $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($entry.'Site Url')
    $Ctx.Credentials = $cred
 
    #Abrufen aller Listen
    $Web = $Ctx.Web
    $Lists = $Web.Lists
    $Ctx.Load($Web)
    $Ctx.Load($Lists)
    $Ctx.ExecuteQuery()
    Write-Host $sites.Count

     Write-Progress -Id 1 -Activity “Fortschritt Sites” -status “Überprüft $b"  -percentComplete ($b / $results.Count*100)
     $b++

   #Durchlaufe durch die Listen der Sites
   ForEach($sitelist in $Lists){ 
  
       
       if(($systemlibs -cnotcontains $sitelist.Title) -and ($sitelist.Hidden -eq $false)){ 
     

       $allItems = @() 
       Write-Host "Aktuelle Liste " $sitelist.Title "der Site" $entry.'Site Title' "mit" $sitelist.ItemCount "Einträgen"
        
                    


           Do{ 
           $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
           $Query.ViewXml= $qCommand 
           $Query.ListItemCollectionPosition = $position
        
           $ListItems = $sitelist.GetItems($Query)
           $Ctx.Load($ListItems)
           $Ctx.ExecuteQuery()
           Write-Host $ListItems.Count
              #Durchlaufe alle Items der Library
              $i = 0
              $i += 1
              Write-Host $i

             
              Foreach($item in $ListItems){
             
               Write-Progress -ParentId 1  -Activity “Fortschritt Listen” -status “Überprüft $i"  -percentComplete ($i / $sitelist.ItemCount*100)
              
                    $i++

                
                  $Versions = $Item.File.Versions
                  $Ctx.Load($Versions) 
                  
                  $Ctx.Load($item.File) 
                   if(!$item.File.Properties.ServerObjectIsNull) {
                  $Ctx.ExecuteQuery()
      
                  $Filename = $item["FileLeafRef"];
   
        
                  Write-host -f Yellow "Total Number of Versions Found in '$Filename' : $($Versions.count)"

                   #Überprüfung und Auswertung ob die Anzahl der Versionen mehr als das Limit sind
                   While($Item.File.Versions.Count -gt $VersionsToKeep)
                   {
                       write-host "Deleting Version:" $Versions[0].VersionLabel
                       $Versions[0].DeleteObject()
                       $Ctx.ExecuteQuery()
               
                       #Laden der Versionen
                       $Ctx.Load($item.File.Versions)
                       $Ctx.ExecuteQuery()  
               
        
                   }     
          
        
         
        # Erhalte die Position der vorausgegangenen Seite
        $position = $ListItems.ListItemCollectionPosition


         #Hinzufügen der aktuellen Collection zur allItems Collection
         $allItems += $ListItems.ListItemCollectionPosition

        
       
      } 
      }
      }Until($position -eq $null)
     } 
    }
   }
 
#}catch {"Abbruch durch einen Fehler: " + $_.Exception.Message}

}else{ Write-Host -f Yellow "Tenant URL fehlt - Bitte eintragen"}


