<#
    Autor: Yannick & Lavan
    Datum: 08.05.2023
    Ort: Gibb Bern
    Beschreibung: Das PowerShell-Skript komprimiert einen vom Benutzer angegebenen Ordner in eine ZIP-Datei. Dabei wird geprüft, 
    ob der Ordner und der Speicherort der ZIP-Datei existieren und ob eine Datei mit demselben Namen bereits vorhanden ist. 
    Nach der Komprimierung zählt das Skript die Anzahl der Dateien im Ordner, 
    öffnet den Speicherort im Windows Explorer und sendet eine Bestätigungsnachricht per E-Mail an den Benutzer.
#>
function GetUniqueFileName($filePath) {
    $fileDirectory = [System.IO.Path]::GetDirectoryName($filePath)
    $fileBaseName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
    $fileExtension = [System.IO.Path]::GetExtension($filePath)
    $counter = 1

    while (Test-Path $filePath) {
        $filePath = Join-Path $fileDirectory ($fileBaseName + "_" + $counter + $fileExtension)
        $counter++
    }

    return $filePath
}

# Frage nach dem Pfad zum Ordner, der komprimiert werden soll
Write-Host "Bitte geben Sie den Pfad zum Ordner ein, der komprimiert werden soll:"
$folderPath = Read-Host

# Überprüft ob der angegebene Pfad existiert
if (Test-Path $folderPath) {
    # Fragt den Speicherort der ZIP-Datei
    Write-Host "Bitte geben Sie den Pfad zum Speicherort der ZIP-Datei ein:"
    $destinationPath = Read-Host

    # Überprüft ob der angegebene Speicherort existiert
    if (Test-Path $destinationPath) {
        # Erstellt den Namen der ZIP-Datei basierend auf dem Ordnernamen
        $zipFileName = [System.IO.Path]::GetFileNameWithoutExtension($folderPath) + ".zip"
        $zipFilePath = Join-Path $destinationPath $zipFileName

        # Überprüft ob eine Datei mit demselben Namen bereits existiert
        if (Test-Path $zipFilePath) {
            Write-Host "Eine Datei mit demselben Namen existiert bereits. Möchten Sie die Datei umbenennen? (j/n)"
            $userInput = Read-Host

            switch ($userInput) {
                "j" { $zipFilePath = GetUniqueFileName($zipFilePath) }
                "J" { $zipFilePath = GetUniqueFileName($zipFilePath) }
                default {
                    Write-Host "Vorgang abgebrochen."
                    return
                }
            }
        }

        # Komprimiert den Ordner in eine ZIP-Datei
        Compress-Archive -Path $folderPath -DestinationPath $zipFilePath

        # Zählt die Anzahl der Dateien im komprimierten Ordner
        $files = Get-ChildItem $folderPath -Recurse -File
        $fileCount = 0
        for ($i = 0; $i -lt $files.Count; $i++) {
            $fileCount++
        }

        Write-Host "Ordner wurde erfolgreich komprimiert und als $zipFilePath gespeichert. Es wurden $fileCount Dateien komprimiert."

        # Öffnet den Speicherort im Windows Explorer
        Invoke-Item $destinationPath

        # Fragt nach der E-Mail-Adresse des Benutzers
        Write-Host "Bitte geben Sie Ihre E-Mail-Adresse ein:"
        $userEmail = Read-Host

        # E-Mail-Adresse und App-Passwort (Nicht veröffentlichen)
        $yourEmail = ""
        $yourPassword = ""

        # Sendet eine Bestätigungsnachricht per E-Mail
        $subject = "Bestaetigung: Ordner wurde erfolgreich komprimiert"
        $body = "Der Ordner '$folderPath' wurde erfolgreich komprimiert und als '$zipFilePath' gespeichert. Es wurden $fileCount Dateien komprimiert."

        $securePassword = ConvertTo-SecureString $yourPassword -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential -ArgumentList $yourEmail, $securePassword

        Send-MailMessage -From $yourEmail -To $userEmail -Subject $subject -Body $body -SmtpServer "smtp.gmail.com" -Port 587 -Credential $credential -UseSsl

        Write-Host "Bestätigungsnachricht wurde erfolgreich an $userEmail gesendet."
    }
    else {
        # Wenn der angegebene Speicherort nicht existiert = Fehlermeldung
        Write-Host "Der angegebene Speicherort ist ungültig oder existiert nicht. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut."
    }
}
else {
    # Wenn der angegebene Pfad nicht existiert = Fehlermeldung
    Write-Host "Der angegebene Pfad ist ungültig oder existiert nicht. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut."
}
