# Path of Disag XML export folder
$SourceFolder = Get-ItemPropertyValue 'HKCU:\Software\DisagOSS' 'xml_directory' 
$SourceFolder = $SourceFolder + "\*"
echo $SourceFolder

$XMLFiles = Get-ChildItem -Path $SourceFolder -Include *.xml
$CultureInfo = New-Object System.Globalization.CultureInfo("de-DE")
foreach ($XMLFile in $XMLFiles)
{
    $ErrorInFile = $false
    $XMLDocument = [XML] (Get-Content -Path $XMLFile)

    # Check Competition Type. Convert only Team competition and League competition
    $CompetitionType = $XMLDocument.SelectSingleNode("result") 
    if (! $CompetitionType.type -like "lc_*"){
        continue
    }

    # Checking Target CSV File
    $CSVDocument = $XMLFile.DirectoryName + '\' + $XMLFile.BaseName +'.txt'
    if([System.IO.File]::Exists($CSVDocument)){
        Remove-Item $CSVDocument
    }

    echo $XMLFile.FullName  
    $Teams = $XMLDocument.SelectNodes("//team") 
    foreach ($Team in $Teams){
        "Team: " + $Team.id.ToString() + ":"

        $Shooters = $Team.ChildNodes
        foreach ($Shooter in $Shooters){
            $ErrorMessage = ""
            $ErrorExists = $false;

            if ($Shooter.remark -eq "")
            {
                $Shooter.fullname + ": " + $Shooter.totalscore

                # Check Passnumber                
                $a = [regex]"[0-9]{8}"
                $b = $a.Match($Shooter.identification) 
                if ($b.Success -eq $false){
                    $ErrorExists = $true;
                    $ErrorInFile = $true;
                    $ErrorMessage = 'Fehler: Passnummer von ' + $Shooter.fullname + ' nicht korrekt! Nur Zahlen und 8 Stellen erlaubt.'
                    $ErrorMessage | out-file $CSVDocument -Append -Encoding ascii
                }
        
                # Get Last valid ShootOff Value
                $ShootOffShots = $Shooter.SelectNodes("shots/shootoff/shot")
                $ShootOffValue = 0
                foreach ($ShootOffShot in $ShootOffShots){
                    if($ShootOffShot.isvalid)
                    {$ShootOffValue = $ShootOffShot.'#text'}
                }
                echo $ErrorExists
                # No Error, add shooter to CSV File
                if ($ErrorExists -eq $false){
                    echo "B"
                    $TotalScore    = [decimal] $Shooter.totalscore
                    $ShootOffValue = [decimal] $ShootOffValue
                    $CSVString = $Shooter.identification + ';'+ $TotalScore.ToString($CultureInfo) + ';' + $ShootOffValue.ToString($CultureInfo)
                    $CSVString | out-file $CSVDocument -Append -Encoding ascii
                }
            }
        }
    
        "--------------------" | out-file $CSVDocument -Append -Encoding ascii
    
        # All shooters with remark.
        foreach ($Shooter in $Shooters){ 
            $ErrorMessage = ""

            # shooter with excluded remark. Example AK etc.
            if ($Shooter.remark -ne ""){
                "(AK): " + $Shooter.fullname + ": " + $Shooter.totalscore

                # Check Passnumber                
                $a = [regex]"[0-9]{8}"
                $b = $a.Match($Shooter.identification) 
                if ($b.Success -eq $false){
                    $ErrorExists = $true;
                    $ErrorInFile = $true;
                    $ErrorMessage = 'Fehler: Passnummer von ' + $Shooter.fullname + ' nicht korrekt! Nur Zahlen und 8 Stellen erlaubt.'
                    $ErrorMessage | out-file $CSVDocument -Append -Encoding ascii
                }

                # No Error, add shooter to CSV File
                if ($ErrorExists -eq $false){
                    $Result = [decimal] $Shooter.totalscore
                    $CSVString = $Shooter.identification + ';'+ $Result.ToString($CultureInfo) + ';'
                    $CSVString | out-file $CSVDocument -Append -Encoding ascii
                }
            }
        }

        "====================" | out-file $CSVDocument -Append -Encoding ascii
    }

    # Check/Create Target Folder
    $TargetFolder = $XMLFile.DirectoryName +'\'+ 'RWK_' + (get-date -uformat "%Y-%m-%d".ToString())    
    if (!(Test-Path $TargetFolder -PathType Container)) {
        $TargetDir = New-Item -ItemType Directory -Force -Path $TargetFolder
    }

    # Move Files to Target Folder        
    $CSVDocument = Get-ChildItem $CSVDocument
    if(Test-Path $CSVDocument -PathType Leaf){           
        #Success
        if($ErrorInFile){ 
            $MoveTo = $TargetFolder + "\" + "Fehler_" + $CSVDocument.Name
            Move-Item -Path $CSVDocument -Destination $MoveTo

            # Move XML to Folder
            $MoveTo = $TargetFolder + "\" + 'Fehler_' + $XMLFile.Name
            Move-Item -Path $XMLFile -Destination $MoveTo
        }
        # Error
        if(!$ErrorInFile){ 
            # Move CSV to Folder   
            $MoveTo = $TargetFolder + "\" + $CSVDocument.Name
            move-Item -Path $CSVDocument -Destination $MoveTo

            # Move XML to Folder
            $MoveTo = $TargetFolder + "\" + $XMLFile.Name
            Move-Item -Path $XMLFile -Destination $MoveTo
        }        
    }
}