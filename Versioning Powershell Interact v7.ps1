# Texte ASCII 3D "VEJA"
Write-Host @"
                                                    
 ___      ___  _______          ___   ________     
|\  \    /  /||\  ___ \        |\  \ |\   __  \                  
\ \  \  /  / /\ \   __/|       \ \  \\ \  \|\  \   
 \ \  \/  / /  \ \  \_|/__   __ \ \  \\ \   __  \  
  \ \    / /    \ \  \_|\ \ |\  \\_\  \\ \  \ \  \ 
   \ \__/ /      \ \_______\\ \________\\ \__\ \__\   
    \|__|/        \|_______| \|________| \|__|\|__|
                                                   

Script PowerShell de Gestion du Versioning SharePoint
____________________________________________________

Ce script permet de gérer le versioning des bibliothèques SharePoint Online, y compris l'activation, la désactivation, et la configuration des versions majeures. 
Il offre également une fonctionnalité de liste des bibliothèques avec leurs statistiques, ainsi que la possibilité de supprimer sélectivement l'historique de versions des fichiers et de vider les corbeilles des sites. 

Version: v7
____________________________________________________
                                                    
"@

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

function Lister-Bibliotheques {
    $bibliothequesDictionary = @{}  # Réinitialiser le dictionnaire

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

    $bibliotheques = $listesEtBibliotheques | Where-Object { ($_.BaseType -eq "DocumentLibrary") -and ($_.hidden -eq $false) -and ($_.Title -ne "Bibliothèque de styles") -and ($_.Title -ne "Style Library") }

    $i = 1
    $versionsActuelles = @{}
    $fileCounts = @{}

    Write-Host "`nBibliothèques disponibles sur $siteName :" -ForegroundColor Cyan
    Write-Host "0: Retour au menu principal"

    foreach ($bibliotheque in $bibliotheques) {
        $bibliothequesDictionary.Add($i, $bibliotheque.Title)

        $contexteBibliotheque = Get-PnPContext
        $liste = $contexteBibliotheque.Web.Lists.GetByTitle($bibliotheque.Title)

        # Charger toutes les propriétés nécessaires en une seule requête
        $contexteBibliotheque.Load($liste)
        $contexteBibliotheque.ExecuteQuery()

        # Utiliser la propriété ItemCount pour obtenir le nombre d'items dans la liste
        $fileCounts[$bibliotheque.Title] = $liste.ItemCount

        $statutVersioning = if ($liste.EnableVersioning) { "Oui" } else { "Non" }
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
                break
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
                    $bibliothequeChoice = $null
                }
            }
        } while ($bibliothequeChoice -eq $null)

        if ($bibliothequeChoice -ne "0") {
            $selectedLibraryTitle = $bibliothequesDictionary[$bibliothequeChoice]
            Afficher-Menu-Bibliotheque $selectedLibraryTitle $versionsActuelles[$selectedLibraryTitle] $fileCounts[$selectedLibraryTitle] $siteName
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
        
        # ... (ajoutez d'autres options de menu ici)
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
    $selectedSiteUrl = $null
    $selectedSiteChosen = $false

    while (-not $selectedSiteChosen) {
        do {
            $siteChoice = $(Write-Host "`nSélectionnez le numéro du site SharePoint à modifier (1 à $($sites.Count)) ,entrez 'quitter' pour quitter le script ou 'versionning' pour modifier le versionning sur tous les sites : " -ForegroundColor yellow -NoNewLine; Read-Host)

            if ($siteChoice -eq "666") {
                Write-Host "I am a man who walks alone" -ForegroundColor DarkGray
                Write-Host "And when I'm walking a dark road" -ForegroundColor DarkGray
                Write-Host "At night or strolling through the park" -ForegroundColor DarkGray
                Write-Host "When the light begins to change" -ForegroundColor DarkGray
                Write-Host "I sometimes feel a little strange" -ForegroundColor DarkGray
                Write-Host "A little anxious when it’s dark" -ForegroundColor DarkGray
                Write-Host "Fear of the dark" -ForegroundColor DarkGray
                Write-Host @"
                          ▄         ▄▄
              ██████████████▄     ,█████████████████████▌
              ███▀▀▀██▀▀▀█▀███▄  ▄██▀▀█████▀▀▀▀▀██▀▀▀▀██▌
              ███╟▓▌█▌▓▓µ▓▓▓▀██████╓▓▓Ç█████▄▓▓▓╗▀U▓▓C██▌
              ███╟▓▌█▌▓▓▓▓╝▓▓╗▀██▀▓▓▓▓▓▓▀███▌▓▓▓▓▓▄▓▓C██▌
              ███╟▓▓█▌▓▓▓▓@▓▓▓▄▀╓▓▓\██╙▓▓╗▀█▌▓▓▌▄╙▓▓▓C██▌
              ███╟▓▌█▌▓▓▓▓▓▓"█▀▓▓▓,▀▀▀▀,▓▓▓╜▌▓▓▌██L▓▓C██▌
              ███╟▓▌█▌▓▓L▄╚▓▓╗█▄▓▓▓▓▓▓▓▓▓▓y█▌▓▓▌██▌▓▓C██▌
           ████████████████▄▓▓▓▀███████████████████╙▓C████████████████████████
            ▀███▄mmmÇ███╓▀███╙▄███▀╔mm▐█]mm▐▌pmmmÇ███`██▌mmmmm╗▐▄τmm╖▀█▌gm╕███
              ▀██▌▓▓▓▓▀▓▓▓╗▀██▀/▀╓▓▓▓▓▐█]▓▓▐▌▓▓▓▓▓▓▀████▌▓▓▀▓▓▓▐██▐▓▓▓╗▀▓▓▌███
               ██▌▓▓▀▓▓▓▀▓▓▐█▄▓▓▓▓▓.▓▓▐█]▓▓▐▌▓▓L▄╙▓▓╗███▌▓▓╥╥╥╥▐██]▓▓▀▓▓▓▓▌███
               ██▌▓▓¬▄╙▄U▓▓▐█▀▄▓▓▓╖^▓▓▐█]▓▓▐▌▓▓L██▄▓▓▓▀█▌▓▓▓▄▄▄▄██]▓▓▐▄╙▓▓▌███
               ██▌▓▓¬███ ▓▓▐`▓▓╩▄▓▓▓▓▓▐█]▓▓▐▌▓▓▓▓▓▓╖▓▓▀▄█╚▓▄███▐██]▓▓▐██╟▓▌███
               ██▌▓▓▐███U▓▓▐█▄▄███╜▓▓▓▐█╘▓▓▐▌▓▓▓▓▓▓▓▓▄████▄▓▓▓▓▐██]╨╨▐██▓▓▌███
               █████████▄╚▓▐████▀██████████████████████▀▀███████████████▄╙▌███
                       ███▄▐██▀                                        ▀██▄███
                        ▀█████⌐                                          ▀████
                          ▀███⌐                                           '▀██
                            ▀█`                                             `▀


"@

                $siteChoice = $null # Pour demander à nouveau le numéro de site
            } elseif ($siteChoice -eq "quitter") {
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
    $selectedSiteUrl = "https://vejafairtradesarl.sharepoint.com/sites/TESTVERSIONNING"
    $siteName = "TESTVERSIONING"
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

function Vider-Seconde-Corbeille-Site {
    try {
        $items = Get-PnPRecycleBinItem
        if (-not $items) {
            Write-Host "La corbeille de deuxième étape du site SharePoint est déjà vide." -ForegroundColor Yellow
            return
        } elseif ($items.Count -eq 0) {
            Write-Host "La corbeille de deuxième étape du site SharePoint est déjà vide." -ForegroundColor Yellow
            return
        }

        $confirmation = Read-Host "Êtes-vous sûr de vouloir vider la corbeille de deuxième étape du site SharePoint ? Entrez 'VIDER' pour confirmer."
        
        if ($confirmation -eq "VIDER") {
            # Vider la corbeille de deuxième étape
            Clear-PnPRecycleBinItem -Force -SecondStageOnly

            Write-Host "Tous les éléments de la corbeille de deuxième étape ont été supprimés." -ForegroundColor Green
        } else {
            Write-Host "Opération annulée. La corbeille de deuxième étape n'a pas été vidée." -ForegroundColor Yellow
        }

    } catch {
        Write-Host "Erreur lors du vidage de la corbeille de deuxième étape du site SharePoint : $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Vider-Corbeille-Site {
    param (
        $siteName
    )

    $batchSize = 100
    $exportFileName = "Vidage-Corbeille-$siteName-$(Get-Date -Format 'yyyyMMddHHmmss').csv"
    $exportFilePath = Join-Path -Path $env:USERPROFILE -ChildPath $exportFileName

    # Avant le premier ajout de données :
    if (-not (Test-Path $exportFilePath)) {
        "Nom du site;Stage;Taille Avant Vidage (Mo);Taille Après Vidage (Mo);Nombre de Fichiers Supprimés" | Out-File -Append -Encoding UTF8 -FilePath $exportFilePath
    }

    $Context = Get-PnPContext
    $RecycleBinItems = $Context.Site.RecycleBin
    $Context.Load($RecycleBinItems)
    $Context.ExecuteQuery()

    $tailleAvantVidage = $RecycleBinItems | Measure-Object -Property Size -Sum | Select-Object -ExpandProperty Sum
    $tailleAvantVidageMo = [math]::Round(($tailleAvantVidage / 1MB), 2)
        
    if ($tailleAvantVidageMo -eq 0) {
        Write-Host "La corbeille du site SharePoint est vide." -ForegroundColor Yellow
    } else {
        $confirmation = Read-Host "Êtes-vous sûr de vouloir vider la corbeille du site $siteName ? Entrez 'VIDER' pour confirmer."
        
        if ($confirmation -eq "VIDER") {
            $itemsToDelete = @($RecycleBinItems)
            $counter = 0

            foreach ($item in $itemsToDelete) {
                $item.DeleteObject()
                $counter++
                
                if ($counter -eq $batchSize) {
                    $Context.ExecuteQuery()
                    $counter = 0
                }
            }

            if ($counter -ne 0) {
                $Context.ExecuteQuery()
            }

            $Context.Load($Context.Site.RecycleBin)
            $Context.ExecuteQuery()

            $tailleApresVidage = $Context.Site.RecycleBin | Measure-Object -Property Size -Sum | Select-Object -ExpandProperty Sum
            $tailleApresVidageMo = [math]::Round(($tailleApresVidage / 1MB), 2)

            # Pour la première corbeille :
            $csvLine1 = "$siteName;Première Corbeille;$tailleAvantVidageMo;$tailleApresVidageMo;$($itemsToDelete.Count)"
            $csvLine1 | Out-File -Append -Encoding UTF8 -FilePath $exportFilePath
            Write-Host "Export effectué avec succès pour la corbeille. Fichier sauvegardé à : $exportFilePath" -ForegroundColor Green
                
            Write-Host "La corbeille du site SharePoint a été vidée avec succès."
            Write-Host "`nNombre de fichiers supprimés : $($itemsToDelete.Count)" -ForegroundColor Cyan
            Write-Host "Taille de la corbeille avant vidage : $($tailleAvantVidageMo) Mo" -ForegroundColor Yellow
            Write-Host "Taille de la corbeille après vidage : $($tailleApresVidageMo) Mo" -ForegroundColor Green
        } else {
            Write-Host "Opération annulée. La corbeille n'a pas été vidée." -ForegroundColor Yellow
        }
    }

    # Demande pour vider la deuxième corbeille, indépendamment de la première corbeille.
    $confirmation2 = Read-Host "Voulez-vous également vider la corbeille de deuxième étape du site $siteName ? (Oui/Non)"

    if ($confirmation2 -eq "Oui") {
        $initialItemsSecondStage = Get-PnPRecycleBinItem
        if ($initialItemsSecondStage.Count -eq 0) {
            Write-Host "La corbeille de deuxième étape du site SharePoint est déjà vide." -ForegroundColor Yellow
            return
        }

        $tailleAvantVidage2 = ($initialItemsSecondStage | Measure-Object -Property Size -Sum).Sum
        $tailleAvantVidageMo2 = [math]::Round(($tailleAvantVidage2 / 1MB), 2)

        Vider-Seconde-Corbeille-Site

        $finalItemsSecondStage = Get-PnPRecycleBinItem
        $tailleApresVidage2 = ($finalItemsSecondStage | Measure-Object -Property Size -Sum).Sum
        $tailleApresVidageMo2 = [math]::Round(($tailleApresVidage2 / 1MB), 2)
            
        $itemsDeleted2 = $initialItemsSecondStage.Count - $finalItemsSecondStage.Count

        # Pour la deuxième corbeille :
        $csvLine2 = "$siteName;Deuxième Corbeille;$tailleAvantVidageMo2;$tailleApresVidageMo2;$itemsDeleted2"
        $csvLine2 | Out-File -Append -Encoding UTF8 -FilePath $exportFilePath
        Write-Host "Export effectué avec succès pour la deuxième corbeille. Fichier sauvegardé à : $exportFilePath" -ForegroundColor Green

        Write-Host "La corbeille de deuxième étape du site SharePoint a été vidée avec succès."
        Write-Host "`nNombre de fichiers supprimés : $($itemsDeleted2)" -ForegroundColor Cyan
        Write-Host "Taille de la corbeille avant vidage : $($tailleAvantVidageMo2) Mo" -ForegroundColor Yellow
        Write-Host "Taille de la corbeille après vidage : $($tailleApresVidageMo2) Mo" -ForegroundColor Green
    } elseif ($confirmation2 -eq "Non") {
        Menu-Principal
    } else {
        Write-Host "Choix non reconnu. Retour au menu principal." -ForegroundColor Red
        Menu-Principal
    }
}

function Supprimer-Historique-Versions-Bibliotheque {
    param (
        [string]$bibliothequeTitle,
        $versionsAConserver,
        $siteName
    )

    $sw = [Diagnostics.Stopwatch]::StartNew()  # Démarrer le chronomètre

    try {
        $smtpServer = "mxa-00848a01.gslb.pphosted.com"  
        $mailFrom = "sharepoint@veja.fr" 
        $mailTo = @("gael.potin@veja.fr", "florent.hautcoeur@veja.fr", "rodney.antoine@veja.fr", "xavier.janin@veja.fr")  

        if ($versionsAConserver -eq "N/A") {
            $versionsAConserver = 1
        }

        $bibliotheque = Get-PnPList -Identity $bibliothequeTitle -ErrorAction Stop
        $confirmation = Read-Host "Êtes-vous sûr de vouloir supprimer l'historique de versions pour la bibliothèque $bibliothequeTitle ? Entrez 'SUPPRIMER' pour confirmer."

        if ($confirmation -eq "SUPPRIMER") {
            $exportFileName = "$($bibliothequeTitle.Replace(" ", "_"))_Suppression_Operations_$((Get-Date).ToString("yyyyMMddHHmm"))_$siteName.csv"
            $exportFilePath = Join-Path -Path $env:USERPROFILE -ChildPath $exportFileName
            $fileStream = [System.IO.StreamWriter]::new($exportFilePath, $false, [System.Text.Encoding]::UTF8)
            $header = "Bibliothèque;Nom du fichier;URL du fichier;Taille du fichier (KB);Nb Versions;Nb Versions avant opérations;Taille des versions (KB);Taille totale du fichier avant opération (KB);Delta;Action effectuée"
            $fileStream.WriteLine($header)

            $files = Get-PnPListItem -List $bibliothequeTitle -PageSize 1000 | Where-Object { $_.FileSystemObjectType -eq "File" }
            $batchSize = 100
            $totalFileCount = $files.Count
            $totalBatches = [math]::Ceiling($totalFileCount / $batchSize)
            $batchCounter = 0
            $totalVersionningBaisses = 0

            while ($batchCounter * $batchSize -lt $totalFileCount) {
                $batchCounter++
                $progress = [math]::Round(($batchCounter / $totalBatches) * 100, 2)
                Write-Progress -Activity "Traitement du lot $batchCounter/$totalBatches en cours..." -Status "$progress% Complété" -PercentComplete $progress
                
                $fichiersTraites = 0
                $versionningBaisses = 0
                $startIndex = ($batchCounter - 1) * $batchSize
                $endIndex = [Math]::Min($startIndex + $batchSize - 1, $totalFileCount - 1)
                $batches = $files[$startIndex..$endIndex]

                foreach ($file in $batches) {
                    $fileName = $file["FileLeafRef"]
                    $fileDir = $file["FileDirRef"]
                    
                    $versions = Get-PnPProperty -ClientObject $file -Property Versions
                
                    if ($null -ne $fileName -and $null -ne $fileDir -and $null -ne $versions) {
                        $result = Process-Versions $file $versions $versionsAConserver $batchCounter $totalBatches $fileStream $siteName $bibliothequeTitle $versionningBaisses $fichiersTraites
                        $versionningBaisses += ($result.VersionningBaisses - $versionningBaisses)
                        $fichiersTraites = $result.FichiersTraites
                    } else {
                        Write-Host "Certaines propriétés sont nulles pour le fichier $fileDir/$fileName. Ignoré." -ForegroundColor Yellow
                    }
                }

                $totalVersionningBaisses += $versionningBaisses

                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()

                Write-Host "Lot $batchCounter/$totalBatches traité - $fichiersTraites fichier(s) - $versionningBaisses versionning baissé(s)" -ForegroundColor Green

                if ($batchCounter -lt $totalBatches) {
                    Write-Host "Temporisation de 10 secondes avant le traitement du lot suivant..." -ForegroundColor Cyan
                    Start-Sleep -Seconds 10
                }
            }

            # Fermer la barre de progression
            Write-Progress -Activity "Traitement des lots terminé" -Completed

            $fileStream.Close()
            $fileStream.Dispose()

            Write-Host "Résultats exportés vers $exportFilePath" -ForegroundColor Green
            Write-Host "Nombre total de versionning baissé: $totalVersionningBaisses" -ForegroundColor Green
            Write-Host ("Mail envoyé à " + ($mailTo -join ", ")) -ForegroundColor Green

        $mailmessage = New-Object system.net.mail.mailmessage 

            foreach ($recipient in $mailTo) {
                $mailmessage.To.add($recipient)
            }

        $smtp = New-Object Net.Mail.SmtpClient($smtpServer) 
        $mailmessage.from = ($mailFrom)
        $mailmessage.To.add($mailTo)
        $mailmessage.Subject = "Résumé de Suppression d'Historique de Versions du site $sitename"
        $mailmessage.IsBodyHtml = $true
        $mailmessage.BodyEncoding = [System.Text.Encoding]::UTF8
        $mailmessage.SubjectEncoding = [System.Text.Encoding]::UTF8

        # Construisez votre corps de mail ici
        $mailBody = "<html><head><meta charset='UTF-8'></head><body>"
        $mailBody += "<h3 style='color: #1976D2;'>Résumé de Suppression d'Historique de Versions de la bibliothèque $bibliothequeTitle du site $sitename</h3>"
        $mailBody += "<p>Nombre de lots traités: <strong style='color: #43A047;'>$totalBatches</strong></p>"
        $mailBody += "<p>Nombre total de fichiers traités: <strong style='color: green;'>$totalFileCount</strong></p>"
        $mailBody += "<p>Nombre total de versionning baissé: <strong style='color: red;'>$totalVersionningBaisses</strong></p>"
        $mailBody += "<p>Résultat final: <strong>Voir fichier joint: </strong>$exportFileName</p>"
        $mailBody += "</body></html>"
        $mailmessage.Body = $mailBody

        $attachment = New-Object System.Net.Mail.Attachment -ArgumentList $exportFilePath
        $mailmessage.Attachments.Add($attachment)

        $smtp.Send($mailmessage)
        $attachment.Dispose()

        } else {
            Write-Host "Opération annulée. L'historique de versions n'a pas été supprimé." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Erreur lors de la suppression de l'historique de versions de la bibliothèque $bibliothequeTitle : $($_.Exception.Message)" -ForegroundColor Red
    } finally {
     $sw.Stop()  # Arrêter le chronomètre
       $elapsedTime = $sw.Elapsed.ToString()
      Write-Host "Durée totale du traitement : $elapsedTime" -ForegroundColor Green
    }
}

function Process-Versions($file, $versions, $versionsAConserver, $batchCounter, $totalBatches, $fileStream, $siteName, $bibliothequeTitle, $versionningBaisses, $fichiersTraites) {
    $fileName = $file["FileLeafRef"]
    $fileDir = $file["FileDirRef"]
    $versionsAvantOperations = $versions.Count

    try {
        $fileBefore = Get-PnPListItem -List $bibliothequeTitle -Id $file.Id
        $FileSizeAvantOperations = [double][Math]::Round($fileBefore["File_x0020_Size"] / 1KB, 2)
    } catch {
        Write-Host "Erreur lors de la récupération de l'élément de la liste pour $fileName dans $bibliothequeTitle. Détails de l'erreur : $_ Ignoré." -ForegroundColor Yellow
        return @{
            VersionningBaisses = $versionningBaisses
            FichiersTraites = $fichiersTraites
        }
    }

    $VersionSizeAvantOperations = [double][Math]::Round(($versions.FieldValues.File_x0020_Size | Measure-Object -Sum).Sum / 1KB, 2)

    if ($versionsAvantOperations -le 1) {
        $TotalSizeAvantOperations = $FileSizeAvantOperations
    } else {
        $TotalSizeAvantOperations = [double]($FileSizeAvantOperations + $VersionSizeAvantOperations)
    }

    $action = "Aucune action nécessaire"
    $TotalSizeApresOperations = $TotalSizeAvantOperations
    $DeltaSize = 0

    # Boucle de vérification et suppression des versions supplémentaires si nécessaire
    while ($versions.Count -gt $versionsAConserver) {
        $delta = $versions.Count - $versionsAConserver  # Recalculer le delta

        $retryLimit = 3
        $retryCount = 0
        $isError = $false

        do {
            try {
                $isError = $false  # Indicateur d'erreur

                $versions | Select-Object -Last $delta | ForEach-Object { $_.DeleteObject() }
                Invoke-PnPQuery
            } catch {
                if ($_.Exception.Message -like "*The collection has not been initialized*") {
                    Write-Host "Erreur d'initialisation de collection pour $fileName dans $bibliothequeTitle. Détails de l'erreur : $_ Ignoré." -ForegroundColor Yellow
                    $isError = $false  # Sortir de la boucle
                } elseif ($_.Exception.Message -like "*has timed out*") {
                    $isError = $true
                    $retryCount++
                    Write-Host "Erreur de délai d'attente dépassé. Tentative $retryCount de $retryLimit..."
                    Start-Sleep -Seconds 10  # Attendre 10 secondes avant de réessayer
                    $versions = Get-PnPProperty -ClientObject $file -Property Versions  # Récupérer à nouveau les versions
                    $delta = $versions.Count - $versionsAConserver  # Recalculer le delta
                } else {
                    Write-Host "Erreur lors de la suppression de l'historique de versions de la bibliothèque $bibliothequeTitle : $_"
                    $isError = $false  # Sortir de la boucle
                }
            } finally {
                $isError = $false  # Assurez-vous que la boucle ne se répète pas indéfiniment
            }
        } while ($isError -and $retryCount -lt $retryLimit)

        $versions = Get-PnPProperty -ClientObject $file -Property Versions  # Récupérer à nouveau les versions
    }

    $VersionSizeApresOperations = [double][Math]::Round(($versions.FieldValues.File_x0020_Size | Measure-Object -Sum).Sum / 1KB, 2)

    if ($versions.Count -le 1) {
        $TotalSizeApresOperations = $FileSizeAvantOperations
    } else {
        $TotalSizeApresOperations = [double]($FileSizeAvantOperations + $VersionSizeApresOperations)
    }

    $DeltaSize = [double][Math]::Round($TotalSizeAvantOperations - $TotalSizeApresOperations, 2)
    $action = "Versionning baissé"

    $TotalSizeApresOperationsStr = $TotalSizeApresOperations.ToString("N2", [cultureinfo]::GetCultureInfo("fr-FR")).Replace(" ", "")
    $VersionSizeAvantOperationsStr = $VersionSizeAvantOperations.ToString("N2", [cultureinfo]::GetCultureInfo("fr-FR")).Replace(" ", "")
    $TotalSizeAvantOperationsStr = $TotalSizeAvantOperations.ToString("N2", [cultureinfo]::GetCultureInfo("fr-FR")).Replace(" ", "")
    $DeltaSizeStr = $DeltaSize.ToString("N2", [cultureinfo]::GetCultureInfo("fr-FR")).Replace(" ", "")

    $csvLine = "$bibliothequeTitle;$fileName;$($file["FileRef"]);$TotalSizeApresOperationsStr;$($versions.Count);$versionsAvantOperations;$VersionSizeAvantOperationsStr;$TotalSizeAvantOperationsStr;$DeltaSizeStr;$action"

    $fileStream.WriteLine($csvLine)
    $fichiersTraites++

    return @{
        VersionningBaisses = $versionningBaisses
        FichiersTraites = $fichiersTraites
    }
}

# Menu principal
function Menu-Principal {
    while ($true) {
        if ($modeDebug -eq $true) {
            Write-Host "`n[Mode Debug : ON]" -ForegroundColor Yellow
        }
        $libraryChoice = Afficher-Menu-Et-Obtenir-Choix @("Activer ou modifier le versioning pour tout le site", "Lister les bibliothèques", "Désactiver le versioning pour tout le site", "Vider la corbeille du site", "Vider la corbeille de deuxième étape du site") "`nVoulez-vous activer ou modifier le versionning pour tout le site $siteName, lister les bibliothèques, désactiver le versioning pour tout le site, vider la corbeille du site ou vider la corbeille de deuxième étape du site ?"

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
            } elseif ($libraryChoice -eq "5") {
                # Vider la corbeille de deuxième étape du site SharePoint
                Vider-Seconde-Corbeille-Site $siteName
            }
        }
    }
}

Menu-Principal