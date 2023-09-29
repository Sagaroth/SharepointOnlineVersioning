# Fonction pour valider une URL SharePoint Online
function Valider-URL-SharePoint([string]$url) {
    if ($url -match "^https:\/\/[a-zA-Z0-9.-]+\.sharepoint\.com\/$") {
        return $true
    } else {
        return $false
    }
}

# Demander à l'utilisateur de saisir l'URL du site admin SharePoint
$adminSiteUrl = $null

# Définir le mode de débogage (true pour activer, false pour désactiver)
$modeDebug = $false

while ($adminSiteUrl -eq $null) {
    $adminSiteUrl = $(Write-Host "Veuillez entrer l'URL du site admin SharePoint : " -ForegroundColor yellow -NoNewLine; Read-Host)

    if (-not (Valider-URL-SharePoint $adminSiteUrl)) {
        Write-Host "URL SharePoint invalide. Veuillez saisir une URL valide." -ForegroundColor Red
        $adminSiteUrl = $null
    }
}

# Tenter de se connecter au service SharePoint
try {
    Connect-SPOService -url $adminSiteUrl
}
catch {
    Write-Host "Erreur lors de la connexion au service SharePoint : $($_.Exception.Message)" -ForegroundColor Red
    Exit
}


# Fonction pour afficher un menu et obtenir une entrée de l'utilisateur avec option de navigation précédente
function Afficher-Menu-Et-Obtenir-Choix([string[]]$options, [string]$prompt) {
    Write-Host $prompt
    for ($i = 0; $i -lt $options.Count; $i++) {
        Write-Host "$($i + 1): $($options[$i])"
    }
    Write-Host "0: Retour en arrière"
    Write-Host "99: Quitter le script" # Ajout de l'option pour quitter le script

    $choix = Read-Host "`nVotre choix"
    return $choix
}

# Fonction pour activer le versioning pour toutes les bibliothèques d'un site SharePoint
function Activer-Versioning-Pour-Toutes-Les-Bibliotheques {

    # Demander le nombre de versions majeures à conserver pour toutes les bibliothèques
    $versionsMajeures = Read-Host "`nCombien de versions majeures voulez-vous conserver pour toutes les bibliothèques (entrez '0' pour désactiver) ?"

    if ($versionsMajeures -eq "0") {
        $choix = Read-Host "Vous avez entré '0', voulez-vous désactiver le versioning ou revenir en arrière ? (entrez 'D' pour désactiver, 'R' pour revenir en arrière)"
        if ($choix -eq "D") {
            Desactiver-Versioning-Pour-Site
            return
        } elseif ($choix -eq "R") {
            return
        } else {
            Write-Host "Choix invalide. Le versioning n'a pas été modifié." -ForegroundColor Yellow
            return
        }
    }

    $listesEtBibliotheques = Get-PnPList

    # Filtrer uniquement les bibliothèques de documents (en excluant "Bibliothèque de styles" ou "Style Library")
    $bibliotheques = $listesEtBibliotheques | Where-Object { ($_.BaseType -eq "DocumentLibrary") -and ($_.hidden -eq $false) -and ($_.Title -ne "Bibliothèque de styles") -and ($_.Title -ne "Style Library") }

    foreach ($bibliotheque in $bibliotheques) {

        if ($($bibliotheque.MajorVersionLimit) -ne $versionsMajeures) {
            Set-PnPList -Identity $bibliotheque.Title -EnableVersioning $true -MajorVersions $versionsMajeures
            Write-Host "Le versionning a été activé pour la bibliothèque $($bibliotheque.Title) avec $versionsMajeures versions majeures." -ForegroundColor Green
        } else {
            Write-Host "Le versioning de la bibliothèque $($bibliotheque.Title) est déjà défini à $versionsMajeures versions majeures. Aucune modification n'est nécessaire." -ForegroundColor Yellow
        }
    }

    Write-Host "Le versionning a été activé pour toutes les bibliothèques du site SharePoint $siteName." -ForegroundColor Green
}

# Fonction pour désactiver le versioning pour tout le site SharePoint
function Desactiver-Versioning-Pour-Site {
    $choixConfirmation = $(Write-Host "`nÊtes-vous sûr de vouloir désactiver le versioning pour tout le site SharePoint ? " -NoNewLine) + $(Write-Host "Oui" -ForegroundColor yellow -NoNewLine) + $(Write-Host "/" -NoNewLine) + $(Write-Host "Non" -ForegroundColor yellow -NoNewLine) + $(Write-Host " " -NoNewLine; Read-Host)

    if ($choixConfirmation -eq "Oui") {
        # Désactiver le versioning pour tout le site SharePoint
        $listesEtBibliotheques = Get-PnPList

        # Filtrer uniquement les bibliothèques de documents (en excluant "Bibliothèque de styles" ou "Style Library")
        $bibliotheques = $listesEtBibliotheques | Where-Object { ($_.BaseType -eq "DocumentLibrary") -and ($_.hidden -eq $false) -and ($_.Title -ne "Bibliothèque de styles") -and ($_.Title -ne "Style Library") }

        foreach ($bibliotheque in $bibliotheques) {
            Set-PnPList -Identity $bibliotheque.Title -EnableVersioning $false
            Write-Host "Le versionning a été désactivé pour la bibliothèque $($bibliotheque.Title)." -ForegroundColor Green
        }

        Write-Host "Le versionning a été désactivé pour toutes les bibliothèques du site SharePoint $siteName." -ForegroundColor Green
    } else {
        Write-Host "Opération annulée." -ForegroundColor Yellow
    }
}

# Fonction pour lister les bibliothèques du site SharePoint avec le nombre de fichiers
function Lister-Bibliotheques {
    $bibliothequesDictionary = @{}  # Réinitialiser le dictionnaire

    # Tenter de récupérer toutes les listes et bibliothèques du site SharePoint sélectionné
    try {
        $listesEtBibliotheques = Get-PnPList -ErrorAction Stop
    }
    catch {
        if ($_.Exception.Message -like "*403*") {
            Write-Host "Erreur : Vous n'avez pas l'autorisation d'accéder aux bibliothèques du site SharePoint $siteName.`n" -ForegroundColor Red
            Disconnect-PnPOnline
            Selectionner-Site-SharePoint
            return
        }
        else {
            Write-Host "Erreur lors de la récupération des bibliothèques : $($_.Exception.Message)" -ForegroundColor Red
            Disconnect-PnPOnline
            Selectionner-Site-SharePoint
            return
        }
    }

    # Filtrer uniquement les bibliothèques de documents (en excluant "Bibliothèque de styles" ou "Style Library")
    $bibliotheques = $listesEtBibliotheques | Where-Object { ($_.BaseType -eq "DocumentLibrary") -and ($_.hidden -eq $false) -and ($_.Title -ne "Bibliothèque de styles") -and ($_.Title -ne "Style Library") }

    $i = 1  # Déclaration de la variable $i et initialisation à 1
    $versionsActuelles = @{}  # Créer un tableau associatif pour stocker les versions actuelles
    $fileCounts = @{}  # Créer un tableau associatif pour stocker le nombre de fichiers

    # Afficher la liste des bibliothèques avec des numéros
    Write-Host "`nBibliothèques disponibles sur $siteName :" -ForegroundColor Cyan

    # Ajouter l'option "0: Retour au menu principal"
    Write-Host "0: Retour au menu principal"

    foreach ($bibliotheque in $bibliotheques) {
        $bibliothequesDictionary.Add($i, $bibliotheque.Title)
        $contexteBibliotheque = Get-PnPContext
        $liste = $contexteBibliotheque.Web.Lists.GetByTitle($bibliotheque.Title)
        $contexteBibliotheque.Load($liste)
        $contexteBibliotheque.ExecuteQuery()
        $versioningActive = $liste.EnableVersioning
        $limiteVersions = $liste.MajorVersionLimit

        # Obtenir tous les éléments de la bibliothèque
        $ListItems = Get-PnPListItem -List $liste.Title -PageSize 500
        $fileCount = $ListItems.Count
        $fileCounts[$bibliotheque.Title] = $fileCount

        $statutVersioning = "Non"
        if ($versioningActive) {
            $statutVersioning = "Oui"
        }

        $versionsActuelles[$bibliotheque.Title] = if ($liste.EnableVersioning) { $liste.MajorVersionLimit } else { "N/A" }
        Write-Host "{$i}: $($bibliotheque.Title), Versioning activé : $statutVersioning, Limite de versions : $($versionsActuelles[$bibliotheque.Title]), Nombre de fichiers : $($fileCounts[$bibliotheque.Title])"
        $i++
    }

    # Sélectionner une bibliothèque en utilisant un numéro
    $bibliothequeChoice = $null

    while ($bibliothequeChoice -eq $null) {
        do {
            $bibliothequeChoice = $(Write-Host "`nSélectionnez le numéro de la bibliothèque à modifier (1 à $($bibliotheques.Count)) ou entrez 'quitter' pour quitter le script: " -ForegroundColor yellow -NoNewLine; Read-Host)

            if ($bibliothequeChoice -eq "0") {
                Write-Host "Retour au menu principal." -ForegroundColor Yellow
                break  # Sortir de la boucle de sélection des bibliothèques pour retourner au menu principal
            } elseif ($bibliothequeChoice -eq "quitter") {
                Write-Host "Au revoir !" -ForegroundColor Yellow
                Disconnect-PnPOnline
                Disconnect-SPOService
                Exit
            } else {
                try {
                    $bibliothequeChoice = [int]$bibliothequeChoice
                    if (-not $bibliothequesDictionary.ContainsKey($bibliothequeChoice)) {
                        Write-Host "Numéro de bibliothèque invalide." -ForegroundColor Yellow
                        $bibliothequeChoice = $null
                    }
                } catch {
                    Write-Host "Entrée invalide. Veuillez saisir un numéro valide ou 'quitter'." -ForegroundColor Yellow
                    $bibliothequeChoice = $null  # Réinitialiser le choix pour redemander
                }
            }
        } while ($bibliothequeChoice -eq $null)

        if ($bibliothequeChoice -ne "0") {
            $selectedLibraryTitle = $bibliothequesDictionary[$bibliothequeChoice]
            Afficher-Menu-Bibliotheque $selectedLibraryTitle $versionsActuelles[$selectedLibraryTitle] $fileCounts[$selectedLibraryTitle] $siteName # Appeler la fonction pour afficher le menu spécifique à la bibliothèque
        }
    }
}

# Fonction pour activer ou modifier le versioning pour toutes les bibliothèques de tous les sites SharePoint
function Modifier-Versioning-Tous-Les-Sites {
    # Demander le nombre de versions majeures à conserver pour toutes les bibliothèques
    $versionsMajeures = Read-Host "`nCombien de versions majeures voulez-vous conserver pour toutes les bibliothèques (entrez '0' pour désactiver) ?"

    if ($versionsMajeures -eq "0") {
        $choix = Read-Host "Vous avez entré '0', voulez-vous désactiver le versioning ou revenir en arrière ? (entrez 'D' pour désactiver, 'R' pour revenir en arrière)"
        if ($choix -eq "D") {
            Desactiver-Versioning-Tous-Les-Sites
            return
        } elseif ($choix -eq "R") {
            return
        } else {
            Write-Host "Choix invalide. Le versioning n'a pas été modifié." -ForegroundColor Yellow
            return
        }
    }

    # Récupérer la liste de tous les sites SharePoint dans votre tenant
    $sites = Get-PnPTenantSite -Detailed

    foreach ($site in $sites) {
        $siteUrl = $site.Url
        Write-Host "Modification du versioning pour le site SharePoint : $siteUrl" -ForegroundColor Cyan

        # Gestion de l'étranglement (throttling) avec réessai
        $RetryCount = 0
        $MaxRetries = 3
        $StopLoop = $false

        do {
            try {
                $siteResult = Get-PnPTenantSite -Url $siteUrl -Detailed

                # Se connecter au site SharePoint
                Connect-PnPOnline -Url $siteUrl -UseWebLogin

                # Récupérer la liste de toutes les bibliothèques de documents du site
                $bibliotheques = Get-PnPList | Where-Object { ($_.BaseType -eq "DocumentLibrary") -and ($_.Hidden -eq $false) -and ($_.Title -ne "Bibliothèque de styles") -and ($_.Title -ne "Style Library") }

                foreach ($bibliotheque in $bibliotheques) {
                    $currentVersionsMajeures = $bibliotheque.EnableVersioningSettings.MajorVersionLimit
                    if ($($bibliotheque.MajorVersionLimit) -ne $versionsMajeures) {
                        Set-PnPList -Identity $bibliotheque.Title -EnableVersioning $true -MajorVersions $versionsMajeures
                        Write-Host "Le versionning a été activé pour la bibliothèque $($bibliotheque.Title) avec $versionsMajeures versions majeures." -ForegroundColor Green
                    } else {
                        Write-Host "Le versionning de la bibliothèque $($bibliotheque.Title) est déjà défini à $versionsMajeures versions majeures. Aucune modification n'est nécessaire." -ForegroundColor Yellow
                    }
                }

                # Déconnecter du site SharePoint actuel
                Disconnect-PnPOnline
                $StopLoop = $true
            }
            catch {
                if ($RetryCount -ge $MaxRetries) {
                    Write-Host "Impossible de compléter l'opération après $MaxRetries réessais pour le site $siteUrl." -ForegroundColor Red
                    $StopLoop = $true
                }
                else {
                    Write-Host "Suite de l'execution du script dans 10 secondes pour le site $siteUrl." -ForegroundColor Yellow
                    Start-Sleep -Seconds 10
                    Connect-PnPOnline -Url $siteUrl -UseWebLogin
                    $RetryCount++
                }
            }
        }
        while (-not $StopLoop)
    }

    Write-Host "Le versioning a été activé pour toutes les bibliothèques de tous les sites SharePoint dans votre tenant." -ForegroundColor Green
}

function Afficher-Menu-Bibliotheque {
    param (
        [string]$bibliothequeTitle,
        $versionsActuelles,
        [int]$fileCounts,
        $siteName
    )

    try {
        # Afficher le menu d'action pour la bibliothèque spécifiée avec le nombre de fichiers
        Write-Host "`nMenu d'action pour la bibliothèque $bibliothequeTitle (Nombre de fichiers : $fileCounts)"
        Write-Host "1: Activer ou modifier le versioning"
        Write-Host "2: Désactiver le versioning"

        # Afficher l'option de suppression de l'historique de versions
        if ($versionsActuelles -eq "N/A") {
            Write-Host "3: Supprimer l'historique de versions des documents"
        } else {
            Write-Host "3: Supprimer l'historique de versions des documents (Limite de versions  : $versionsActuelles)"
        }
        Write-Host "0: Retour au menu précédent"

        # Demander à l'utilisateur de choisir une action
        $actionChoice = $(Write-Host "`nSélectionnez une action (1, 2, 3, 0, etc.) : " -ForegroundColor yellow -NoNewLine; Read-Host)

        if ($actionChoice -eq "1") {
            # Code pour activer le versioning de la bibliothèque
            Activer-Versioning-Bibliotheque $bibliothequeTitle $versionsActuelles
        } elseif ($actionChoice -eq "2") {
            # Code pour désactiver le versioning de la bibliothèque
            Desactiver-Versioning-Bibliotheque $bibliothequeTitle
        } elseif ($actionChoice -eq "3") {
            # Code pour supprimer l'historique de versions des documents de la bibliothèque
            Supprimer-Historique-Versions-Bibliotheque $bibliothequeTitle $versionsActuelles $siteName
        } elseif ($actionChoice -eq "0") {
            # Retourner au menu précédent (dans ce cas, cela sortira de cette fonction)
            return
        } else {
            # Gérer d'autres choix d'action ici, si nécessaire
        }
    }
    catch {
        Write-Host "Erreur lors de l'affichage du menu pour la bibliothèque $bibliothequeTitle : $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Fonction pour activer le versioning d'une bibliothèque avec une limite de versions
function Activer-Versioning-Bibliotheque {
    param (
        [string]$bibliothequeTitle,
        $versionsActuelles
    )

    try {
        # Demander le nombre de versions majeures
        $limiteVersions = Read-Host "Entrez le nombre de versions majeures (1 à 50000) :"

        if ([int]$limiteVersions -ge 1 -and [int]$limiteVersions -le 50000) {
            # Récupérer la bibliothèque spécifiée
            $bibliotheque = Get-PnPList -Identity $bibliothequeTitle -ErrorAction Stop

            if ($versionsActuelles -ne $limiteVersions) {
                # Activer le versioning sur la bibliothèque avec la limite de versions spécifiée
                Set-PnPList -Identity $bibliotheque -EnableVersioning $true -MajorVersions $limiteVersions
                Write-Host "Le versioning a été activé pour la bibliothèque $bibliothequeTitle avec une limite de versions de $limiteVersions." -ForegroundColor Green
            } else {
                Write-Host "Le versioning de la bibliothèque $bibliothequeTitle est déjà défini à $limiteVersions versions majeures. Aucune modification n'est nécessaire." -ForegroundColor Yellow
            }
        } else {
            Write-Host "Le nombre de versions majeures doit être compris entre 1 et 50000. Activation du versioning annulée." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Erreur lors de l'activation du versioning pour la bibliothèque $bibliothequeTitle : $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Fonction pour désactiver le versioning d'une bibliothèque
function Desactiver-Versioning-Bibliotheque {
    param (
        [string]$bibliothequeTitle
    )

    try {
        # Récupérer la bibliothèque spécifiée
        $bibliotheque = Get-PnPList -Identity $bibliothequeTitle -ErrorAction Stop

        # Désactiver le versioning pour la bibliothèque
        Set-PnPList -Identity $bibliotheque -EnableVersioning $false

        Write-Host "Le versioning a été désactivé pour la bibliothèque $bibliothequeTitle." -ForegroundColor Green
    }
    catch {
        Write-Host "Erreur lors de la désactivation du versioning pour la bibliothèque $bibliothequeTitle : $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Déclarez $siteName en tant que variable globale
$global:siteName = $null

# Fonction pour sélectionner un site SharePoint
function Selectionner-Site-SharePoint {
    $sites = Get-SPOSite -Limit All

    if ($sites.Count -eq 0) {
        Write-Host "Aucun site SharePoint accessible trouvé." -ForegroundColor Yellow
        Disconnect-SPOService
        Disconnect-PnPOnline
        Exit
    }

    # Créer un dictionnaire pour stocker les numéros et les URL des sites SharePoint
    $sitesDictionary = @{}
    $sortedSites = $sites | Sort-Object Url

    # Afficher la liste des sites SharePoint accessibles avec des numéros (triés par ordre alphabétique)
    Write-Host "Sites SharePoint accessibles :" -ForegroundColor Cyan
    for ($i = 0; $i -lt $sortedSites.Count; $i++) {
        $site = $sortedSites[$i]
        $sitesDictionary.Add($i + 1, $site.Url)
        Write-Host "$($i + 1): $($site.Url)"
    }

    # Sélectionner un site SharePoint en utilisant un numéro ou en entrant "quitter" pour quitter le script
    $siteChoice = $null
    $libraryChoice = $null

    $selectedSiteUrl = $null  # Initialisez la variable ici
    $selectedSiteChosen = $false  # Ajout d'une variable pour suivre si un site a été choisi

    while (-not $selectedSiteChosen) {
        do {
            $siteChoice = $(Write-Host "`nSélectionnez le numéro du site SharePoint à modifier (1 à $($sites.Count)) ,entrez 'quitter' pour quitter le script ou 'versionning' pour modifier le versionning sur tous les sites : " -ForegroundColor yellow -NoNewLine; Read-Host)

            if ($siteChoice -eq "quitter") {
                Write-Host "Au revoir !" -ForegroundColor Yellow
                if ($selectedSiteUrl -ne $null) {
                    Disconnect-PnPOnline
                }
                Disconnect-SPOService
                Exit
            } elseif ($siteChoice -eq "versionning") {
                # Appeler la fonction de versionning global
                Modifier-Versioning-Tous-Les-Sites
                return
            } else {
                try {
                    $siteChoice = [int]$siteChoice
                    if (-not $sitesDictionary.ContainsKey($siteChoice)) {
                        Write-Host "Numéro de site SharePoint invalide." -ForegroundColor Yellow
                        $siteChoice = $null
                    }
                } catch {
                    Write-Host "Entrée invalide. Veuillez saisir un numéro valide ou 'quitter'." -ForegroundColor Yellow
                    $siteChoice = $null  # Réinitialiser le choix pour redemander
                }
            }
        } while ($siteChoice -eq $null)

        $selectedSiteUrl = $sitesDictionary[$siteChoice]

        # Tenter de se connecter au site SharePoint sélectionné
        try {
            Connect-PnPOnline -Url $selectedSiteUrl -UseWebLogin
            $selectedSiteChosen = $true  # Un site a été choisi, mettez la variable à vrai
        }
        catch {
            if ($_.Exception.Message -like "*403*") {
                Write-Host "Erreur : Vous n'avez pas l'autorisation d'accéder au site SharePoint." -ForegroundColor Red
                if ($selectedSiteUrl -ne $null) {
                    Disconnect-PnPOnline
                }
                $selectedSiteUrl = $null
            }
            else {
                Write-Host "Erreur lors de la connexion au site SharePoint : $($_.Exception.Message)" -ForegroundColor Red
                if ($selectedSiteUrl -ne $null) {
                    Disconnect-PnPOnline
                }
                $selectedSiteUrl = $null
            }
        }

        if ($selectedSiteUrl -ne $null) {
            # Mettez à jour la variable globale $siteName avec le nom du site SharePoint
            $global:siteName = $selectedSiteUrl -split '/' | Select-Object -Last 1
        }
    }

    Write-Host "Vous êtes maintenant connecté au site SharePoint $siteName." -ForegroundColor Green
}

if ($modeDebug -eq $true) {
    $selectedSiteUrl = ""
    $siteName = ""
    Connect-PnPOnline -Url $selectedSiteUrl -UseWebLogin
} else {
    # Sélectionner un site SharePoint lorsque le mode débug est désactivé
    Selectionner-Site-SharePoint
}

# Fonction pour convertir une taille en octets en une chaîne lisible (ko, Mo, Go, etc.)
function Convertir-TailleEnOctets {
    param (
        [double]$tailleEnOctets
    )

    $unites = "octets", "ko", "Mo", "Go", "To"
    $index = 0

    while ($tailleEnOctets -ge 1024 -and $index -lt $unites.Length - 1) {
        $tailleEnOctets /= 1024
        $index++
    }

    return "{0:N2} {1}" -f $tailleEnOctets, $unites[$index]
}

# Fonction pour vider la corbeille d'un site SharePoint
function Vider-Corbeille-Site {
    param (
        $siteName
    )

    try {
        # Obtenir le contexte SharePoint
        $Context = Get-PnPContext

        # Récupérer la corbeille du site SharePoint
        $RecycleBinItems = $Context.Site.RecycleBin
        $Context.Load($RecycleBinItems)
        $Context.ExecuteQuery()

        if ($RecycleBinItems.Count -eq 0) {
            Write-Host "La corbeille du site SharePoint est vide." -ForegroundColor Yellow
        } else {
            # Demander une confirmation à l'utilisateur
            $confirmation = Read-Host "Êtes-vous sûr de vouloir vider la corbeille du site $siteName ? Entrez 'VIDER' pour confirmer."

            if ($confirmation -eq "VIDER") {
                # Récupérer la taille de la corbeille avant vidage
                $tailleAvantVidage = 0
                $RecycleBinItems | ForEach-Object { $tailleAvantVidage += $_.Size }

                # Créer un tableau pour stocker les éléments à supprimer
                $elementsASupprimer = @()

                # Ajouter les éléments à supprimer au tableau
                $RecycleBinItems | ForEach-Object {
                    $elementsASupprimer += $_
                }

                # Supprimer les éléments du tableau
                foreach ($element in $elementsASupprimer) {
                    $element.DeleteObject()
                }
                $Context.ExecuteQuery()

                # Récupérer à nouveau la corbeille après vidage
                $RecycleBinItemsApresVidage = $Context.Site.RecycleBin
                $Context.Load($RecycleBinItemsApresVidage)
                $Context.ExecuteQuery()
                
                # Récupérer la taille de la corbeille après vidage
                $tailleApresVidage = 0
                $RecycleBinItemsApresVidage | ForEach-Object { $tailleApresVidage += $_.Size }

                # Compter le nombre de fichiers supprimés
                $nombreFichiersSupprimes = $elementsASupprimer.Count

                Write-Host "La corbeille du site SharePoint a été vidée avec succès."
                Write-Host "`nNombre de fichiers supprimés : $nombreFichiersSupprimes" -ForegroundColor Cyan
                Write-Host "Taille de la corbeille avant vidage : $(Convertir-TailleEnOctets $tailleAvantVidage)" -ForegroundColor Yellow
                Write-Host "Taille de la corbeille après vidage : $(Convertir-TailleEnOctets $tailleApresVidage)" -ForegroundColor Green

                # Créer un objet PSObject pour stocker les données
                $vidageDetails = New-Object PSObject -Property @{
                    "Nom du site" = $siteName
                    "Taille Avant Vidage" = (Convertir-TailleEnOctets $tailleAvantVidage)
                    "Taille Après Vidage" = (Convertir-TailleEnOctets $tailleApresVidage)
                    "Nombre de Fichiers Supprimés" = $nombreFichiersSupprimes
                }

                # Exporter l'objet PSObject vers Excel
                $exportFileName = "Vidage-Corbeille-$siteName-$(Get-Date -Format 'yyyyMMddHHmmss').xlsx"
                $exportFilePath = Join-Path -Path $env:USERPROFILE -ChildPath $exportFileName
                $vidageDetails | Export-Excel -Path $exportFilePath -WorksheetName "Vidage Details" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow

                Write-Host "`nDétails de l'opération de vidage exportés vers : $exportFilePath" -ForegroundColor Green
            } else {
                Write-Host "Opération annulée. La corbeille n'a pas été vidée." -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Host "Erreur lors du vidage de la corbeille du site SharePoint : $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Supprimer-Historique-Versions-Bibliotheque {
    param (
        [string]$bibliothequeTitle,
        $versionsAConserver,
        $siteName
    )

    try {
        if ($versionsAConserver -eq "N/A") {
            $versionsAConserver = 1
        }

        $bibliotheque = Get-PnPList -Identity $bibliothequeTitle -ErrorAction Stop
        $suppressionOperations = @()
        $confirmation = Read-Host "Êtes-vous sûr de vouloir supprimer l'historique de versions pour la bibliothèque $bibliothequeTitle ? Entrez 'SUPPRIMER' pour confirmer."

        if ($confirmation -eq "SUPPRIMER") {
            $files = Get-PnPListItem -List $bibliothequeTitle -PageSize 1000 | Where-Object { $_.FileSystemObjectType -eq "File" }

            # Divisez les fichiers en paquets de taille raisonnable
            $batchSize = 100  # Choisissez la taille du paquet appropriée
            $batches = [System.Collections.ArrayList]@()
            $batch = [System.Collections.ArrayList]@()

            foreach ($file in $files) {
                $batch.Add($file) | Out-Null
                if ($batch.Count -eq $batchSize) {
                    $batches.Add($batch.Clone()) | Out-Null
                    $batch.Clear()
                }
            }

            if ($batch.Count -gt 0) {
                $batches.Add($batch.Clone()) | Out-Null
            }

            $batchCounter = 1

            foreach ($batch in $batches) {
                $StopLoop = $false
                $RetryCount = 0
                $MaxRetries = 3
                do {
                    try {
                        foreach ($file in $batch) {
                            $fileName = $file["FileLeafRef"]
                            $fileDir = $file["FileDirRef"]
                            $VersionSizeinKB = 0
                            $versions = Get-PnPProperty -ClientObject $file -Property Versions
                            $libraryName = $bibliotheque.Title
                            $versionsAvantOperations = $versions.Count
                            $VersionSizeAvantOperations = 0
                            foreach ($version in $versions) {
                                $VersionSizeAvantOperations += [Math]::Round(($version.FieldValues.File_x0020_Size / 1KB), 2)
                            }
                            # Validation des propriétés avant utilisation
                            if ($null -ne $fileName -and $null -ne $fileDir -and $null -ne $versions -and $null -ne $libraryName) {
                                if ($versions.Count -gt $versionsAConserver) {
                                    $delta = $versions.Count - $versionsAConserver
                                    for ($i = $versions.Count - 1; $i -ge $versionsAConserver; $i--) {
                                        $versions[$i].DeleteObject() | Out-Null
                                    }
                                    Invoke-PnPQuery | Out-Null
                                    $versions = Get-PnPProperty -ClientObject $file -Property Versions
                                    $Status = "Versionning baissé"
                                    Write-Host "($batchCounter/$($batches.Count)) L'historique de versions du fichier $fileDir/$fileName a été réduit à $versionsAConserver versions (Versions supprimées : $delta)." -ForegroundColor Green
                                } elseif ($versions.Count -eq 1) {
                                    $Status = "Pas de versionning"
                                    Write-Host "($batchCounter/$($batches.Count)) Pas de versionning pour le fichier $fileDir/$fileName." -ForegroundColor Yellow
                                } elseif ($versions.Count -eq $versionsAConserver) {
                                    $Status = "Aucune action nécessaire"
                                    Write-Host "($batchCounter/$($batches.Count)) L'historique de versions du fichier $fileDir/$fileName est déjà égal à $versionsAConserver versions. Aucune action nécessaire." -ForegroundColor Yellow
                                } elseif ($versions.Count -lt $versionsAConserver) {
                                    $Status = "Aucune action nécessaire"
                                    Write-Host "($batchCounter/$($batches.Count)) L'historique de versions du fichier $fileDir/$fileName est en dessous de $versionsAConserver versions. Aucune action nécessaire." -ForegroundColor Yellow
                                } else {
                                    $Status = "Aucune action nécessaire"
                                    Write-Host "($batchCounter/$($batches.Count)) Aucune action nécessaire pour le fichier $fileDir/$fileName." -ForegroundColor Cyan
                                }

                                $VersionSizeinKB = 0
                                foreach ($version in $versions) {
                                    $VersionSizeinKB += [Math]::Round(($version.FieldValues.File_x0020_Size / 1KB), 2)
                                }

                                $TotalFileSizeKB = [Math]::Round(($file["File_x0020_Size"] / 1KB + $VersionSizeinKB), 2)
                                $versionsSizeKB = [Math]::Round(($VersionSizeinKB), 2)
                                if ($versionsSizeKB -eq [Math]::Round(($file["File_x0020_Size"] / 1KB), 2)) {
                                    $versionsSizeKB = "Pas de versionning"
                                }

                                $suppressionOperations += [PSCustomObject]@{
                                    "Bibliothèque" = $bibliothequeTitle
                                    "Nom du fichier" = $fileName
                                    "URL du fichier" = $file["FileRef"]
                                    "Taille du fichier (KB)" = [Math]::Round(($file["File_x0020_Size"] / 1KB), 2)
                                    "Nb Versions" = if ($versionsSizeKB -eq "Pas de versionning") { "Pas de versionning" } else { $versions.Count }
                                    "Nb Versions avant opérations" = $versionsAvantOperations
                                    "Taille des versions (KB)" = $versionsSizeKB
                                    "Taille des versions avant Opération" = $VersionSizeAvantOperations
                                    "Taille totale du fichier (KB)" = [Math]::Round(($VersionSizeinKB), 2)
                                    "Taille totale du fichier avant opération (KB)" = if ($VersionSizeinKB -eq 0) { [Math]::Round(($file["File_x0020_Size"] / 1KB), 2) } else { [Math]::Round(($VersionSizeAvantOperations), 2) }
                                    "Action effectuée" = $Status
                                }
                            }
                            else {
                                Write-Host "($batchCounter/$($batches.Count)) Certaines propriétés sont nulles pour le fichier $fileDir/$fileName. Ignoré." -ForegroundColor Yellow
                            }
                        }

                        $StopLoop = $true
                    } catch {
                        if ($RetryCount -ge $MaxRetries) {
                            Write-Host "Impossible de compléter l'opération après $MaxRetries réessais pour la bibliothèque $bibliothequeTitle." -ForegroundColor Red
                            $StopLoop = $true
                        } else {
                            Write-Host "Suite de l'exécution du script dans 5 secondes après une erreur : $($_.Exception.Message)" -ForegroundColor Yellow
                            Start-Sleep -Seconds 5
                            $RetryCount++
                        }
                    }
                } while (-not $StopLoop)

                $batchCounter++
                
                # Ajoutez une temporisation de 10 secondes après chaque lot
                if ($batchCounter -le $batches.Count) {
                    Write-Host "Temporisation de 10 secondes avant le traitement du lot suivant..." -ForegroundColor Cyan
                    Start-Sleep -Seconds 10
                }
            }

            # Ajoutez la somme des colonnes
            $sumRow = @{
                "Bibliothèque" = "Somme"
                "Nom du fichier" = ""
                "URL du fichier" = ""
                "Taille du fichier (KB)" = ""
                "Nb Versions" = ""
                "Nb Versions avant opérations" = ""
                "Taille des versions (KB)" = ""
                "Taille des versions avant Opération" = ""
                "Taille totale du fichier (KB)" = ($suppressionOperations | Measure-Object -Property "Taille totale du fichier (KB)" -Sum).Sum
                "Taille totale du fichier avant opération (KB)" = ($suppressionOperations | Measure-Object -Property "Taille totale du fichier avant opération (KB)" -Sum).Sum
                "Action effectuée" = ""
            }

            $suppressionOperations += New-Object PSObject -Property $sumRow

            $exportFileName = "Suppression-Historique-$siteName-$bibliothequeTitle-$(Get-Date -Format 'yyyyMMddHHmmss').xlsx"
            $exportFilePath = Join-Path -Path $env:USERPROFILE -ChildPath $exportFileName

            $suppressionOperations | Export-Excel -Path $exportFilePath -WorksheetName "Opérations $sitename" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow -CellStyleSB {
                param($worksheet)
        
                $worksheet.Cells["A1:K1"].Style.HorizontalAlignment = "Center"
                $worksheet.Cells["A2:C$($suppressionOperations.Count + 1)"].Style.HorizontalAlignment = "Left"
                $worksheet.Cells["D2:K$($suppressionOperations.Count + 1)"].Style.HorizontalAlignment = "Center"
            }

            Write-Host "Opérations de suppression exportées vers : $exportFilePath" -ForegroundColor Green

        } else {
            Write-Host "Opération annulée. L'historique de versions n'a pas été supprimé." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Erreur lors de la suppression de l'historique de versions de la bibliothèque $bibliothequeTitle : $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Menu principal
function Menu-Principal {
    while ($true) {
        if ($modeDebug -eq $true) {
            Write-Host "`n[Mode Debug : ON]" -ForegroundColor Yellow
        }
        $libraryChoice = Afficher-Menu-Et-Obtenir-Choix @("Activer ou modifier le versioning pour tout le site", "Lister les bibliothèques", "Désactiver le versioning pour tout le site", "Vider la corbeille du site") "`nVoulez-vous activer ou modifier le versionning pour tout le site $siteName, lister les bibliothèques, désactiver le versioning pour tout le site ou vider la corbeille du site ?"

        if ($libraryChoice -eq "0") {
            # Retourner au menu de sélection de site SharePoint
            if ($modeDebug -eq $false) {
                if ($connectedToSite -eq $true) {
                    Disconnect-PnPOnline  # Déconnectez uniquement si une connexion a été établie
                    $connectedToSite = $false  # Réinitialisez la variable
                }
                $selectedSiteUrl = $null
                Write-Host "Retour au menu de sélection de site.`n" -ForegroundColor Yellow
                Selectionner-Site-SharePoint
            }
        } elseif ($libraryChoice -eq "99") {
            # Quitter le script
            Write-Host "Au revoir !" -ForegroundColor Yellow
            if ($connectedToSite -eq $true) {
                Disconnect-PnPOnline
            }
            Disconnect-SPOService

            Exit
        } else {
            if ($libraryChoice -eq "1") {
                # Activer le versioning pour tout le site SharePoint
                Activer-Versioning-Pour-Toutes-Les-Bibliotheques $versionsActuelles
            } elseif ($libraryChoice -eq "2") {
                # Lister les bibliothèques du site SharePoint
                Lister-Bibliotheques
            } elseif ($libraryChoice -eq "3") {
                # Désactiver le versioning pour tout le site SharePoint
                Desactiver-Versioning-Pour-Site
            } elseif ($libraryChoice -eq "4") {
                # Vider la corbeille du site SharePoint
                Vider-Corbeille-Site $siteName
            }
        }
    }
}

Menu-Principal
