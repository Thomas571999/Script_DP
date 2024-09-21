# Import des fonctions Win32 et du style XAML pour manipuler les fenêtres d'apparition utilisateur
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@


Add-Type -AssemblyName PresentationFramework

# Première fenêtre XAML
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Programme d'export des données Wat.Erp"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        SizeToContent="WidthAndHeight"
        WindowStyle="SingleBorderWindow"
        Background="White">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" TextAlignment="Center" FontWeight="Bold" Foreground="Red" FontSize="16" TextWrapping="Wrap" Margin="0,0,0,10">
            /!\ ATTENTION /!\
        </TextBlock>

        <TextBlock Grid.Row="1" TextAlignment="Center" TextWrapping="Wrap" Margin="0,10,0,20" FontSize="14">
            • Ce programme va analyser les modifications apportées depuis le dernier export MODULARIS.<LineBreak/><LineBreak/>
            • Assurez-vous d'avoir sauvegardé les précédents fichiers générés par ce programme vers un autre dossier avant de continuer. En cas de problème vous pourrez malgré tout les retrouver au format backup dans le dossier de travail que vous allez sélectionner.<LineBreak/><LineBreak/>
            Si vous n'êtes pas prêt, quittez et relancez le programme plus tard.
        </TextBlock>

        <!-- ComboBox pour le choix entre 1, 2 et 3 -->
        <StackPanel Grid.Row="2" Orientation="Vertical" HorizontalAlignment="Center" Margin="0,20,0,0">
            <TextBlock Text="Veuillez sélectionner l'étape TAB pour laquelle sont prévus les fichiers :" Margin="0,0,0,5" TextAlignment="Center" FontSize="14"/>
            <ComboBox Name="OptionComboBox" Width="200" Height="30" SelectedIndex="-1" HorizontalContentAlignment="Center">
                <ComboBoxItem Content="TAB_1" />
                <ComboBoxItem Content="TAB_2" />
                <ComboBoxItem Content="TAB_3" />
            </ComboBox>
        </StackPanel>

        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,20,0,0">
            <Button Name="ContinueButton" Width="150" Height="40" Margin="5" Content="Continuer" IsDefault="True" Background="#BFFFBF" IsEnabled="False"/>
            <Button Name="QuitButton" Width="150" Height="40" Margin="5" Content="Quitter" IsCancel="True" Background="#FFBFBF"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Récupération des éléments ComboBox et boutons
$optionComboBox = $window.FindName("OptionComboBox")
$continueButton = $window.FindName("ContinueButton")
$quitButton = $window.FindName("QuitButton")

# Variable pour stocker le choix de l'utilisateur
$global:TAB = $null

# Activation du bouton Continuer seulement si une option est sélectionnée
$optionComboBox.add_SelectionChanged({
    if ($optionComboBox.SelectedIndex -ne -1) {
        $continueButton.IsEnabled = $true
    } else {
        $continueButton.IsEnabled = $false
    }
})

# Actions sur les boutons
$continueButton.Add_Click({
    $global:TAB = $optionComboBox.SelectedItem.Content
    $window.DialogResult = $true
    $window.Close()
})
$quitButton.Add_Click({
    $window.DialogResult = $false
    $window.Close()
})

# Affichage de la fenêtre
$result = $window.ShowDialog()

# Mettre la fenêtre principale au premier plan
[void][Win32]::SetForegroundWindow([System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle)

# Traitement en fonction du choix de l'utilisateur
if ($result -eq $true) {
    Write-Output "TAB Sélectionné : $TAB"
    Write-Output "Sélectionnez le dossier de sauvegarde vers lequel le programme doit exporter les données"
} else {
    Write-Output "L'exécution du programme a été annulée par l'utilisateur, il s'est arrêté."
    exit
}


# Fenêtre de sélection du dossier dans lequel le script va travailler et sauvegarder les fichiers créés
function Select-FolderDialog {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Sélectionnez le dossier de sauvegarde vers lequel le programme doit exporter les données"
    $folderBrowser.ShowNewFolderButton = $false

    $result = $folderBrowser.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    } else {
        return $null
    }
}

$selectedFolder = Select-FolderDialog

if (-not $selectedFolder) {
    Write-Output "Aucun dossier de travail n'a été sélectionné, le programme s'est arrêté."
    Exit
}

Write-Output "Dossier de travail sélectionné : $selectedFolder"


# Définition des paths des fichiers à traiter
$filesToCheck = @(
    "CSV_ajouts.csv",
    "CSV_suppressions.csv",
    "CSV_modifications.csv",
    "CSV_5_communes.csv",
    "CSV_modifications_propres.csv"
    "Abonnes_Complet.csv"
)

# Initialisation d'une variable pour savoir si un fichier a été trouvé
$filesFound = $false

# Fonction pour renommer et déplacer un fichier
function Rename-And-Move-File($filePath, $backupFolderPath) {
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
    $extension = [System.IO.Path]::GetExtension($filePath)
    $newFileName = $fileName + ".bakup" + $extension
    $newFilePath = Join-Path -Path $backupFolderPath -ChildPath $newFileName

    Move-Item -Path $filePath -Destination $newFilePath
    Write-Output "Fichier $filePath renommé et déplacé vers $newFilePath"
}

# Vérifier et traiter chaque fichier
foreach ($file in $filesToCheck) {
    $filePath = Join-Path -Path $selectedFolder -ChildPath $file

    if (Test-Path $filePath) {
        # Si c'est le premier fichier trouvé, créer le dossier backup avec la date et l'heure actuelles
        if (-Not $filesFound) {
            $backupFolderName = "backup_" + (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
            $backupFolderPath = Join-Path -Path $selectedFolder -ChildPath $backupFolderName
            New-Item -ItemType Directory -Path $backupFolderPath
            $filesFound = $true
        }

        # Renommer et déplacer le fichier
        Rename-And-Move-File -filePath $filePath -backupFolderPath $backupFolderPath
    }
}

# Vérifier s'il n'y avait aucun fichier à déplacer
if (-Not $filesFound) {
    Write-Output "Il n'y a aucun fichier à backup."
}



# Définition des paths des fichiers à supprimer, s'ils existent déjà
$deleted_outputAddedPath = Join-Path -Path $selectedFolder -ChildPath "CSV_ajouts.csv"
$deleted_outputRemovedPath = Join-Path -Path $selectedFolder -ChildPath "CSV_suppressions.csv"
$deleted_outputModifiedPath = Join-Path -Path $selectedFolder -ChildPath "CSV_modifications.csv"
$deleted_output5Communes = Join-Path -Path $selectedFolder -ChildPath "CSV_5_communes.csv"
$deleted_outputprochaine = Join-Path -Path $selectedFolder -ChildPath "CSV_modifications_propres.csv"
$deleted_outputmisajour = Join-Path -Path $selectedFolder -ChildPath "Abonnes_Complet.csv"

# Suppression des fichiers temporaires, s'ils existent déjà
if (-Not(Test-Path $deleted_outputAddedPath)) {
    Write-Output "Il n'y a pas de fichier CSV_ajouts.csv à supprimer"
} else {
    Remove-Item $deleted_outputAddedPath
}
if (-Not(Test-Path -Path $deleted_outputRemovedPath)) {
    Write-Output "Il n'y a pas de fichier CSV_suppressions.csv à supprimer"
} else {
    Remove-Item -Path $deleted_outputRemovedPath
}
if (-Not(Test-Path -Path $deleted_outputModifiedPath)) {
    Write-Output "Il n'y a pas de fichier CSV_modifications.csv à supprimer"
} else {
    Remove-Item -Path $deleted_outputModifiedPath
}
if (-Not(Test-Path -Path $deleted_output5Communes)) {
    Write-Output "Il n'y a pas de fichier CSV_5_communes.csv à supprimer"
} else {
    Remove-Item -Path $deleted_output5Communes
}
if (-Not(Test-Path -Path $deleted_outputprochaine)) {
    Write-Output "Il n'y a pas de fichier CSV_modifications_propres.csv à supprimer"
} else {
    Remove-Item -Path $deleted_outputprochaine
}

# Forcer la sélection de l'utilisateur uniquement sur des fichiers aux extensions CSV
function Select-File {
    param (
        [string]$title = "Sélectionnez un fichier",
        [string]$filter = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
    )

    Add-Type -AssemblyName PresentationFramework

    $openFileDialog = New-Object Microsoft.Win32.OpenFileDialog
    $openFileDialog.Title = $title
    $openFileDialog.Filter = $filter

    $result = $openFileDialog.ShowDialog()

    if ($result -eq $true) {
        return $openFileDialog.FileName
    } else {
        return $null
    }
}

# Afficher le premier message de sélection du précédent fichier d'export (pour permettre la comparaison des données et conclure aux modifications faites)
$title = "Sélection du précédent fichier d'export"
$message = "Dans la prochaine fenêtre, sélectionnez l'ANCIEN fichier d'export abonnés :"
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$title"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        SizeToContent="WidthAndHeight"
        WindowStyle="SingleBorderWindow"
        Background="White">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" TextAlignment="Center" TextWrapping="Wrap" Margin="0,10,0,20" FontSize="14">
            $message
        </TextBlock>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,20,0,0">
            <Button Name="OKButton" Width="100" Height="30" Margin="5" Content="OK" IsDefault="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader] $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$OKButton = $window.FindName("OKButton")
$OKButton.Add_Click({
    $window.DialogResult = $true
    $window.Close()
})
$result = $window.ShowDialog()
[void][Win32]::SetForegroundWindow([System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle)

$oldCsvPath = Select-File -title "SÉLECTIONNEZ L'ANCIEN FICHIER D'EXPORT ABONNÉS"

if ($result -eq $true) {
    Write-Output "Fichier sélectionné comme base de l'ancien export : $oldCsvPath"
} else {
    Write-Output "Aucun fichier n'a été sélectionné. Le programme va s'arrêter."
    Start-Sleep -Seconds 2
}

if (-not $oldCsvPath) {
    Show-MessageBox -message "Aucun fichier n'a été sélectionné. Le programme va s'arrêter." -title "Erreur : aucun fichier sélectionné"
    Exit
}

# Afficher le second message de sélection du fichier SAVIGNE + HOMMES
$title = "Sélection de l'export SAVIGNE + HOMMES"
$message = "Dans la prochaine fenêtre, sélectionnez le fichier d'export SAVIGNE + HOMMES :"
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$title"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        SizeToContent="WidthAndHeight"
        WindowStyle="SingleBorderWindow"
        Background="White">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" TextAlignment="Center" TextWrapping="Wrap" Margin="0,10,0,20" FontSize="14">
            $message
        </TextBlock>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,20,0,0">
            <Button Name="OKButton" Width="100" Height="30" Margin="5" Content="OK" IsDefault="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader] $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$OKButton = $window.FindName("OKButton")
$OKButton.Add_Click({
    $window.DialogResult = $true
    $window.Close()
})
$result = $window.ShowDialog()
[void][Win32]::SetForegroundWindow([System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle)

$CSV_Savigne_Hommes = Select-File -title "SÉLECTIONNEZ LE FICHIER D'EXPORT ABONNÉS SAVIGNE + HOMMES"

if ($result -eq $true) {
    Write-Output "Fichier sélectionné pour les abonnés SAVIGNE + HOMMES : $CSV_Savigne_Hommes"
} else {
    Write-Output "Aucun fichier n'a été sélectionné. Le programme va s'arrêter."
}

if (-not $CSV_Savigne_Hommes) {
    Show-MessageBox -message "Aucun fichier n'a été sélectionné. Le programme va s'arrêter." -title "Erreur : aucun fichier sélectionné"
    Exit
}

# Afficher le troisième message de sélection de l'export AVRILLE + CLERE + MAZIERES
$title = "Sélection de l'export AVRILLE + CLERE + MAZIERES"
$message = "Dans la prochaine fenêtre, sélectionnez le fichier d'export AVRILLE + CLERE + MAZIERES :"
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$title"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        SizeToContent="WidthAndHeight"
        WindowStyle="SingleBorderWindow"
        Background="White">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" TextAlignment="Center" TextWrapping="Wrap" Margin="0,10,0,20" FontSize="14">
            $message
        </TextBlock>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,20,0,0">
            <Button Name="OKButton" Width="100" Height="30" Margin="5" Content="OK" IsDefault="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader] $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$OKButton = $window.FindName("OKButton")
$OKButton.Add_Click({
    $window.DialogResult = $true
    $window.Close()
})
$result = $window.ShowDialog()
[void][Win32]::SetForegroundWindow([System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle)

$CSV_Avrille_Clere_Mazieres = Select-File -title "SÉLECTIONNEZ LE FICHIER D'EXPORT ABONNÉS AVRILLE + CLERE + MAZIERES"

if ($result -eq $true) {
    Write-Output "Fichier sélectionné pour les abonnés AVRILLE + CLERE + MAZIERES : $CSV_Avrille_Clere_Mazieres"
} else {
    Write-Output "Aucun fichier n'a été sélectionné. Le programme va s'arrêter."
}

if (-not $CSV_Avrille_Clere_Mazieres) {
    Show-MessageBox -message "Aucun fichier n'a été sélectionné. Le programme va s'arrêter." -title "Erreur : aucun fichier sélectionné"
    Exit
}



# Afficher le quatrième message de sélection du listing client et le menu de sélection de fichier
$title = "Sélection du fichier propre ABONNÉS_COMPLET"
$message = "Dans la prochaine fenêtre, sélectionnez le fichier propre ABONNÉS_COMPLET :"
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$title"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        SizeToContent="WidthAndHeight"
        WindowStyle="SingleBorderWindow"
        Background="White">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" TextAlignment="Center" TextWrapping="Wrap" Margin="0,10,0,20" FontSize="14">
            $message
        </TextBlock>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,20,0,0">
            <Button Name="OKButton" Width="100" Height="30" Margin="5" Content="OK" IsDefault="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader] $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$OKButton = $window.FindName("OKButton")
$OKButton.Add_Click({
    $window.DialogResult = $true
    $window.Close()
})
$result = $window.ShowDialog()
[void][Win32]::SetForegroundWindow([System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle)

$AbonnesCompletCsvPath = Select-File -title "SÉLECTIONNEZ LE FICHIER ABONNÉS_COMPLET"

if ($result -eq $true) {
    Write-Output "Fichier sélectionné comme base propre Abonnés_Complet : $AbonnesCompletCsvPath"
} else {
    Write-Output "Aucun fichier n'a été sélectionné. Le programme va s'arrêter."
    Start-Sleep -Seconds 2
}

if (-not $AbonnesCompletCsvPath) {
    Show-MessageBox -message "Aucun fichier n'a été sélectionné. Le programme va s'arrêter." -title "Erreur : aucun fichier sélectionné"
    Exit
}




# Fonction pour afficher un timer basé sur la progression
function Show-Progress {
    param (
        [int]$currentStep,      # Étape actuelle
        [int]$totalSteps,       # Nombre total d'étapes
        [string]$taskName       # Nom de la tâche
    )

    # Calcul du pourcentage de progression
    $progressPercent = [math]::Round(($currentStep / $totalSteps) * 100, 2)
    $elapsedTime = (Get-Date) - $global:startTime
    $elapsedSeconds = $elapsedTime.TotalSeconds

    # Estimation du temps restant
    if ($currentStep -gt 0) {
        $estimatedTotalTime = ($elapsedSeconds / $currentStep) * $totalSteps
        $remainingTime = [math]::Round($estimatedTotalTime - $elapsedSeconds, 2)
    } else {
        $remainingTime = "Calcul en cours..."
    }

    # Convertir le temps restant en minutes et secondes si supérieur à 60 secondes
    if ($remainingTime -is [double] -and $remainingTime -ge 60) {
        $minutes = [math]::Floor($remainingTime / 60)
        $seconds = [math]::Round($remainingTime % 60)
        $formattedTime = "$minutes minute(s) et $seconds seconde(s)"
    } else {
        $formattedTime = "$remainingTime seconde(s)"
    }

    # Affichage de la progression
    Write-Host "${taskName}: ${progressPercent}% complet - Temps restant estimé: $formattedTime." -NoNewline
    Write-Host "`r" -NoNewline  # Efface la ligne pour la rafraîchir
}


# Début du timer global pour mesurer le temps total
$global:startTime = Get-Date

# Fusion des fichiers avec progression dynamique
$CSV_5_communes = Join-Path -Path $selectedFolder -ChildPath "CSV_5_communes.csv"

# Lire les contenus de fichier
$savigneHommesData = Get-Content -Path $CSV_Savigne_Hommes
$avrilleClereMazieresData = Get-Content -Path $CSV_Avrille_Clere_Mazieres | Select-Object -Skip 1

$totalLines = $savigneHommesData.Count + $avrilleClereMazieresData.Count
$currentStep = 0

# Combiner les fichiers avec suivi de progression
$combinedData = foreach ($line in $savigneHommesData + $avrilleClereMazieresData) {
    $currentStep++
    
    # Effacement de la ligne précédente avant d'afficher la progression
    Write-Host ("`r" + " " * 80 + "`r") -NoNewline  # Efface la ligne en cours

    Show-Progress -currentStep $currentStep -totalSteps $totalLines -taskName "Fusion des fichiers"
    $line
}

# Exporter les données combinées dans un fichier CSV
$combinedData | Set-Content -Path $CSV_5_communes

Write-Host "Fusion des fichiers terminée."


# Importer à nouveau le fichier combiné
$data55 = Import-Csv -Path $CSV_5_communes -Delimiter ';' -Encoding UTF8

# Ajouter la colonne NUMERO_INCREMENT avec des valeurs incrémentées
$startNumber = 13748
$totalLines = $data55.Count
$currentStep = 0

$modifiedData = foreach ($row in $data55) {
    if ($null -ne $row) {
        $currentStep++
        Show-Progress -currentStep $currentStep -totalSteps $totalLines -taskName "Ajout des NUMERO_INCREMENT"
        
        $row | Add-Member -MemberType NoteProperty -Name NUMERO_INCREMENT -Value $startNumber -Force
        $startNumber++
        $row  # Renvoie l'objet modifié pour le stockage dans $modifiedData
    }
}

# Exporter les données modifiées dans le fichier CSV avec le délimiteur ';'
if ($modifiedData) {
    $modifiedData | Export-Csv -Path $CSV_5_communes -NoTypeInformation -Delimiter ';' -Encoding UTF8
    Write-Host "Les fichiers ont été combinés et la colonne NUMERO_INCREMENT a été ajoutée avec succès dans $CSV_5_communes"
} else {
    Write-Host "Aucune donnée modifiée à exporter."
}


# Fonction pour convertir les dates en format uniforme
function Convert-ToUniformDateFormat {
    param (
        [string]$dateString
    )
    $dateFormats = @(
        'dd/MM/yyyy',
        'MM/dd/yyyy',
        'yyyy-MM-dd'
    )
    
    foreach ($format in $dateFormats) {
        try {
            return [datetime]::ParseExact($dateString, $format, $null).ToString('dd/MM/yyyy')
        } catch {
            continue
        }
    }
    return $dateString
}

# Comparaison avec progression dynamique














# Définition des chemins des fichiers CSV créés
$outputAddedPath = Join-Path -Path $selectedFolder -ChildPath "CSV_ajouts.csv"
$outputRemovedPath = Join-Path -Path $selectedFolder -ChildPath "CSV_suppressions.csv"
$outputModifiedPath = Join-Path -Path $selectedFolder -ChildPath "CSV_modifications.csv"
$outputUpdatedPath = Join-Path -Path $selectedFolder -ChildPath "CSV_modifications_propres.csv"
#$AbonnesCompletCsvPath = Join-Path -Path $selectedFolder -ChildPath "Abonnes_Complet.csv"

# Lire les fichiers CSV
$oldData = Import-Csv -Path $oldCsvPath -Delimiter ';' -Encoding UTF8
$newData = Import-Csv -Path $CSV_5_communes -Delimiter ';' -Encoding UTF8
$AbonnesCompletData = Import-Csv -Path $AbonnesCompletCsvPath -Delimiter ';' -Encoding UTF8

# Créer des dictionnaires pour comparaison
$oldDataDict = @{}
$newDataDict = @{}
$AbonnesCompletDict = @{}

foreach ($row in $oldData) {
    $compositeKey = "$($row.REFAB)_$($row.NCONTAB)"
    $oldDataDict[$compositeKey] = $row
}

foreach ($row in $newData) {
    $compositeKey = "$($row.REFAB)_$($row.NCONTAB)"
    $newDataDict[$compositeKey] = $row
}

foreach ($row in $AbonnesCompletData) {
    $compositeKey = "$($row.REFAB)_$($row.NCONTAB)"
    $AbonnesCompletDict[$compositeKey] = $row
}

$addedRows = @()
$removedRows = @()
$modifiedRows = @()
$updatedData = @()

# Obtenir l'ordre des colonnes depuis $AbonnesCompletData
$columnOrder = $AbonnesCompletData[0].PSObject.Properties.Name





$totalLines = $newDataDict.Count  # Calculer le nombre total de lignes pour comparaison
$currentStep = 0  # Initialiser l'étape actuelle

foreach ($compositeKey in $newDataDict.Keys) {
    $currentStep++  # Incrémentation de l'étape

    Show-Progress -currentStep $currentStep -totalSteps $totalLines -taskName "Comparaison des fichiers"

    if ($oldDataDict.ContainsKey($compositeKey)) {
        $newRow = $newDataDict[$compositeKey]
        $oldRow = $oldDataDict[$compositeKey]
        $abonnesRow = $AbonnesCompletDict[$compositeKey]
        $rowUpdated = $false
        $updatedRow = @{}

        foreach ($property in $newRow.PSObject.Properties) {
            $propertyName = $property.Name
            if ($propertyName -eq 'NUMERO_INCREMENT') { continue }

            $newValue = Convert-ToUniformDateFormat $newRow.$propertyName
            $oldValue = Convert-ToUniformDateFormat $oldRow.$propertyName

            # Si une modification est détectée, on la signale et on stocke la modification
            if ($newValue -ne $oldValue) {
                $modifiedRows += [pscustomobject]@{
                    REFAB = $newRow.REFAB
                    NCONTAB = $newRow.NCONTAB
                    Noms_colonne_modifiee = $propertyName
                    Ancienne_valeur = $oldValue
                    Nouvelle_valeur = $newValue
                }

                # On enregistre la valeur modifiée dans la ligne mise à jour
                $updatedRow[$propertyName] = $newValue
                $rowUpdated = $true
            }
        }

        # Si la ligne a été modifiée, on la complète avec les autres colonnes depuis AbonnesCompletCsvPath
        if ($rowUpdated) {
            foreach ($property in $abonnesRow.PSObject.Properties) {
                $propertyName = $property.Name

                # Si la colonne n'a pas été modifiée, on prend la valeur de $AbonnesCompletCsvPath
                if (-not $updatedRow.ContainsKey($propertyName)) {
                    $updatedRow[$propertyName] = $abonnesRow.$propertyName
                }
            }

            # Construire la ligne mise à jour dans l'ordre des colonnes défini par $columnOrder
            $orderedRow = New-Object PSObject
            foreach ($col in $columnOrder) {
                $orderedRow | Add-Member -MemberType NoteProperty -Name $col -Value $updatedRow[$col]
            }

            # Ajouter la ligne ordonnée dans les données mises à jour
            $updatedData += $orderedRow
        }
    } else {
        $addedRows += $newDataDict[$compositeKey]
    }
}

# Initialiser les variables pour la progression
$totalSteps = $oldDataDict.Keys.Count + $addedRows.Count + $removedRows.Count + $modifiedRows.Count + $abonnesCompletData.Count
$currentStep = 0

# Vérifier les suppressions
foreach ($compositeKey in $oldDataDict.Keys) {
    $currentStep++
    Show-Progress -currentStep $currentStep -totalSteps $totalSteps -taskName "Vérification des suppressions"
    
    if (-not $newDataDict.ContainsKey($compositeKey)) {
        $removedRows += $oldDataDict[$compositeKey]
    }
}

# Export des résultats des ajouts
if ($addedRows.Count -gt 0) {
    $currentStep++
    Show-Progress -currentStep $currentStep -totalSteps $totalSteps -taskName "Exportation des ajouts"
    
    $addedRows | Export-Csv -Path $outputAddedPath -NoTypeInformation -Delimiter ';' -Encoding UTF8
}

# Export des résultats des suppressions
if ($removedRows.Count -gt 0) {
    $currentStep++
    Show-Progress -currentStep $currentStep -totalSteps $totalSteps -taskName "Exportation des suppressions"
    
    $removedRows | Export-Csv -Path $outputRemovedPath -NoTypeInformation -Delimiter ';' -Encoding UTF8
}

# Export des résultats des modifications
if ($modifiedRows.Count -gt 0) {
    $currentStep++
    Show-Progress -currentStep $currentStep -totalSteps $totalSteps -taskName "Exportation des modifications"
    
    $modifiedRows | Export-Csv -Path $outputModifiedPath -NoTypeInformation -Delimiter ';' -Encoding UTF8
}

# Exporter les données mises à jour dans un nouveau fichier CSV pour la prochaine extraction
if ($updatedData.Count -gt 0) {
    $currentStep++
    Show-Progress -currentStep $currentStep -totalSteps $totalSteps -taskName "Exportation des données mises à jour"
    
    $updatedData | Export-Csv -Path $outputUpdatedPath -NoTypeInformation -Delimiter ';' -Encoding UTF8
    Write-Host "Les données mises à jour ont été exportées dans $outputUpdatedPath"
} else {
    Write-Host "Aucune donnée mise à jour pour la prochaine extraction."
}

Write-Host "Comparaison terminée. Les fichiers de résultats ont été exportés."

# Lire les données du fichier Abonnes complet
$abonnesCompletData = Import-Csv -Path $AbonnesCompletCsvPath -Delimiter ';' -Encoding UTF8

# Lire les fichiers des ajouts, suppressions et modifications
if ($addedRows.Count -gt 0) {
    $addedData = Import-Csv -Path $outputAddedPath -Delimiter ';' -Encoding UTF8
}
$removedData = Import-Csv -Path $outputRemovedPath -Delimiter ';' -Encoding UTF8
$modifiedData = Import-Csv -Path $outputUpdatedPath -Delimiter ';' -Encoding UTF8

# Créer un dictionnaire des lignes à supprimer pour une recherche rapide
$removedDict = @{}
foreach ($row in $removedData) {
    $compositeKey = "$($row.REFAB)_$($row.NCONTAB)"
    $removedDict[$compositeKey] = $true
}

# Créer un dictionnaire des lignes modifiées pour une recherche rapide
$modifiedDict = @{}
foreach ($row in $modifiedData) {
    $compositeKey = "$($row.REFAB)_$($row.NCONTAB)"
    $modifiedDict[$compositeKey] = $row
}

# Préparation des données mises à jour
$finalData = @()

# Parcourir les abonnés existants et traiter les suppressions et modifications
foreach ($abonnesRow in $abonnesCompletData) {
    $currentStep++
    Show-Progress -currentStep $currentStep -totalSteps $totalSteps -taskName "Traitement des abonnés existants"
    
    $compositeKey = "$($abonnesRow.REFAB)_$($abonnesRow.NCONTAB)"

    # Si la ligne fait partie des suppressions, on la saute
    if ($removedDict.ContainsKey($compositeKey)) {
        continue
    }

    # Si la ligne fait partie des modifications, on la remplace par la ligne modifiée
    if ($modifiedDict.ContainsKey($compositeKey)) {
        $finalData += $modifiedDict[$compositeKey]
    } else {
        # Sinon, on conserve la ligne telle quelle
        $finalData += $abonnesRow
    }
}

# Ajouter les nouvelles lignes des ajouts
foreach ($addedRow in $addedData) {
    $currentStep++
    Show-Progress -currentStep $currentStep -totalSteps $totalSteps -taskName "Ajout des nouvelles lignes"
    
    $finalData += $addedRow
}

# Exporter le fichier final
$AbonnesCompletMisAJourPath = Join-Path -Path $selectedFolder -ChildPath "Abonnes_Complet.csv"
$finalData | Export-Csv -Path $AbonnesCompletMisAJourPath -NoTypeInformation -Delimiter ';' -Encoding UTF8

Write-Host "La synthèse finale a été exportée dans $AbonnesCompletMisAJourPath"




##########




# Chemins des fichiers CSV temporaires créés à l'étape précédente
$cheminBaseInitial = $AbonnesCompletMisAJourPath

# Chemins d'enregistrement des fichiers CSV finaux
$cheminBaseNouveau = Join-Path -Path $selectedFolder -ChildPath "listing_complet.csv"

# Liste des types de voies et de numéros à détecter
$typesVoie = @("rue", "rte", "imp", "pl", "za", "place", "pl.", "avenue", "allée", "allee", "all", "chemin", "route", "boulevard", "impasse", "place", "quai", "passage", "square", "promenade", "lotissement", "lieu-dit", "moulin", "hameau", "zac", "lot")
$typesNumeros = @("a", "b", "c", "d", "e", "t", "bis", "ter", "quater")

# Importer le premier fichier CSV à traiter
$data = Import-Csv -Path $cheminBaseInitial -Delimiter ';' -Encoding UTF8
$nouvellesDonnees = @()

foreach ($ligne in $data) {
    $adresse_facturation = $ligne.ADRCOMPT
    $increment = $ligne.NUMERO_INCREMENT

    # Initialisation de nouvelles variables pour faire un découpage des adresses
    $numeroRue = ""
    $typeRue = ""
    $nomRue = ""
    $complement = ""

    # Découpage des adresses en utilisant des expressions régulières
    if ($adresse_facturation -match '^(\d+)\s*(bis|ter|quater|[a-z])?\s+(.+)$') {
        $numeroRue = $matches[1]
        $complement = $matches[2]
        $nomRue = $matches[3]
    } else {
        $nomRue = $adresse_facturation
    }

    # Diviser le nom de rue par mots
    $mots = $nomRue -split '\s+'

    # Vérifier si le premier mot est un type de voie
    $premierMot = $mots[0].ToLower()

    if ($typesVoie -contains $premierMot) {
        $typeRue = $mots[0]
        $nomRue = ($mots[1..($mots.Length - 1)]) -join ' '
    }

    # Créer un nouvel objet avec les données formatées
    $nouvelleLigne = [PSCustomObject] @{
        'Numero Increment' = $increment
        'Adresse' = $adresse_facturation
        'Numero Rue' = $numeroRue
        'Type Rue' = $typeRue
        'Nom Rue' = $nomRue
        'Complement' = $complement
    }

    $nouvellesDonnees += $nouvelleLigne
}

# Exporter ces données traitées vers le fichier CSV final
$nouvellesDonnees | Export-Csv -Path $cheminBaseNouveau -Delimiter ';' -NoTypeInformation -Encoding UTF8



###########


# Chemin du fichier CSV source (version UTF-8)
$sourceFile = $AbonnesCompletMisAJourPath

# Chemin du fichier CSV de destination
$destinationFile = Join-Path -Path $selectedFolder -ChildPath "consommations_complet.csv"

# Importer le fichier CSV source
$data = Import-Csv -Path $sourceFile -Delimiter ';'-Encoding UTF8

# Afficher les en-têtes du fichier CSV pour vérifier les noms des colonnes
Write-Host "En-têtes du fichier CSV source :"
$data[0] | Get-Member -MemberType NoteProperty | ForEach-Object { Write-Host $_.Name }

# Afficher les premières lignes du fichier CSV pour vérifier le contenu
Write-Host "`nPremières lignes du fichier CSV source :"
$data | Select-Object -First 5 | Format-Table -AutoSize

# Créer un dictionnaire pour suivre les compteurs d'identifiants par code postal
$idCounters = @{}

# Créer une liste pour stocker les nouvelles lignes
$transformedData = @()

# Traiter chaque ligne du fichier CSV
foreach ($row in $data) {
    #Write-Host "Traitement de la ligne avec REFAB: $($row.REFAB)"

    # Obtenir le code postal
    $codePostal = $row.CPCOMPT

    # Initialiser le compteur pour ce code postal s'il n'existe pas
    if (-not $idCounters.ContainsKey($codePostal)) {
        $idCounters[$codePostal] = 0
    }

    # Boucler sur les colonnes de consommation, date et index
    for ($i = 1; $i -le 5; $i++) {
        $consumptionColumn = "CONSNM$i"
        $dateColumn = "DATM$i"
        $indexColumn = "INDM$i"

        # Vérifier que l'index n'est pas "0" et que les autres colonnes ne sont pas vides
        if ($row.$indexColumn -ne '0' -and $row.$consumptionColumn -ne '' -and $row.$dateColumn -ne '' -and $row.$indexColumn -ne '') {
            #Write-Host "Consommation ($consumptionColumn): $($row.$consumptionColumn)"
            #Write-Host "Date ($dateColumn): $($row.$dateColumn)"
            #Write-Host "Index ($indexColumn): $($row.$indexColumn)"
            
            # Incrémenter le compteur pour ce code postal
            $idCounters[$codePostal]++
            $id = "{0}{1:D5}" -f $codePostal, $idCounters[$codePostal]
            
            # Ajouter la ligne transformée avec l'identifiant unique
            $transformedData += [pscustomobject]@{
                REFAB = $row.REFAB
                Consommation = $row.$consumptionColumn
                Date = $row.$dateColumn
                Index = $row.$indexColumn
                Identifiant = $id
            }
        }
    }
}

# Vérifier si $transformedData contient des données
if ($transformedData.Count -eq 0) {
    Write-Host "Aucune donnée à écrire dans le fichier de sortie."
} else {
    # Exporter les données transformées dans un nouveau fichier CSV
    $transformedData | Export-Csv -Path $destinationFile -NoTypeInformation -Delimiter ';' -Encoding UTF8
    Write-Host "Les données ont été transformées et sauvegardées dans $destinationFile"
}




###########



$dossierinutile = Join-Path -Path $selectedFolder -ChildPath "Fichiers_Pour_Verifier"
$dossierutile = Join-Path -Path $selectedFolder -ChildPath "Fichiers_A_Utiliser"
$dossierprochaintab = Join-Path -Path $selectedFolder -ChildPath "Fichiers_Pour_Prochain_TAB"
New-Item -ItemType Directory -Path $dossierinutile
New-Item -ItemType Directory -Path $dossierutile
New-Item -ItemType Directory -Path $dossierprochaintab

Move-Item -Path $CSV_5_communes -Destination "$dossierprochaintab\Fichier_5_communes_extraction_brute_$TAB.csv"
Copy-Item -Path $AbonnesCompletMisAJourPath -Destination "$dossierprochaintab\Fichier_Abonnes_Complet_$TAB.csv"
Move-Item -Path $AbonnesCompletMisAJourPath -Destination $dossierutile
Move-Item -Path $cheminBaseNouveau -Destination $dossierutile
Move-Item -Path $destinationFile -Destination $dossierutile

# Définir les chemins des fichiers
$paths = @($outputRemovedPath, $outputAddedPath, $outputModifiedPath, $outputUpdatedPath)

# Déplacer les fichiers s'ils existent
foreach ($path in $paths) {
    if (Test-Path $path) {
        Move-Item -Path $path -Destination $dossierinutile
    }
}




Start-Sleep -Seconds 2
Write-Host "Appuyez sur n'importe quelle touche pour terminer ..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
