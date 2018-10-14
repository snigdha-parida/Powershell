Clear-Host
Start-Transcript

$notes = new-object -comobject Lotus.NotesSession
$notes.initialize()
$strDominoDir = 'C:\Users\snigdha.parida\Documents\Powershell\DB\realestate.nsf'
$DomDatabase = $notes.GetDatabase( '', $strDominoDir, 1 )
#-------------------
write-Host $DomDatabase.Title ": Database Now Open for business" 
$domView = $DomDatabase.GetView('House Type Category') 

# Show number of documents
$DomNumOfDocs = $DomView.AllEntries.Count
Write-Host "Num of Docs : " $DomNumOfDocs


# Get First Document in the View
$i = 1;
$DomDocument = $DomView.GetFirstDocument()
while ($DomDocument -ne $null) {

    Write-Host "---------------- Document $i -----------------"

    $items = $DomDocument.Items
    foreach ($item in $items) 
    {
        #if($item.Name -eq "sampleImage" -and $item.type -eq 1)
        if($item.Values -ne $null -and $item.Name -eq "sampleImage")
        {
            $item
        }
    }    


$i+=1;
$DomDocument = $DomView.GetNextDocument($DomDocument)
}



