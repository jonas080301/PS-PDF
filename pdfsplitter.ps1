Import-Module PSWritePDF 
Add-Type -AssemblyName System.Windows.Forms

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = $pwd } #Initialize Filebrowser with default location current working dir

$null = $FileBrowser.ShowDialog() #Open Filebrowser

$file = $FileBrowser.FileName #get FileName of selected file



$date = Get-Date -Format "dd-MM-yyyy" #current date, to set as folder name

if($file){ #if selected file is not null
    $path = (get-item $file ).Directory.FullName #get path of file
    try{
        New-Item -Path "$path\$date" -ItemType directory -ErrorAction Stop #try to create new folder at file location (named after the current date)
    }
    catch{
        [System.Windows.Forms.MessageBox]::Show("Es existiert bereits ein Ordner mit dem heutigen Datum!","Ordner exsitiert bereits!",0) #catch error
    }



    Split-PDF -FilePath "$file" -OutputFolder "$path\$date" #split the PDF

    $childItems = Get-ChildItem -Path "$path\$date" #get the result of the PDF split

    for ($i=0; $i -lt $childItems.length; $i++){ #loop over the new files
        $usedItem = $childItems[$i] 
        $pdfText = Convert-PDFToText -FilePath "$path\$date\$usedItem" #read pdf


        $name = $pdfText.Split([Environment]::NewLine)[2] + "-" + $pdfText.Split([Environment]::NewLine)[8] + ".pdf" #find the name + bookingnumber to set as filename

        $name = $name -replace '\s','' #replace empty characters
        $name = $name -replace 'Name:','' #replace header 'name'
        $name = $name -replace 'Buchungsnummer:','' #replace header 'buchungsnummer'
        $name = $name -replace [Environment]::NewLine,'' #replace linebreaks

        Write-Host $name

        Rename-Item -Path "$path\$date\$usedItem" -NewName "$name" #rename the pdf
    }
}
else{
    [System.Windows.Forms.MessageBox]::Show("Keine Datei ausgewählt!", "Datei fehlt!" , 0) #catch error
}

