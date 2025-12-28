# Vérificateur DNS avancé pour Google Sheets

![License MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green)
![Runtime](https://img.shields.io/badge/Google%20Apps%20Script-V8-green)
![Author](https://img.shields.io/badge/Auteur-Fabrice%20Faucheux-orange)

## Description

Ce projet est un outil d'audit DNS automatisé intégré à Google Sheets. Il permet de vérifier en masse la santé technique d'une liste de domaines via l'API Google DNS (`dns.google`). Le script est conçu pour être résilient, contourner les problèmes de cache DNS locaux et fournir des rapports détaillés sur les enregistrements critiques (MX, SPF, DMARC, A).

Il s'adresse aux administrateurs systèmes, aux équipes de délivrabilité email et aux développeurs souhaitant surveiller des configurations de domaines sans recourir à des outils externes coûteux.

## Fonctionnalités clés

* **Vérification Multi-protocoles** : Analyse des enregistrements MX, A, SPF et DMARC.
* **Interprétation intelligente** :
    * Identification automatique des fournisseurs de messagerie (Google Workspace, Office 365, etc.).
    * Lecture des politiques de sécurité (SPF `softfail`/`hardfail`, DMARC `p=none`/`quarantine`/`reject`).
* **Architecture robuste** :
    * Utilisation du moteur V8 de Google Apps Script.
    * Gestion d'erreurs (Try/Catch) par domaine pour ne pas bloquer l'exécution globale.
    * Système "Anti-Cache" via paramètre aléatoire dans les requêtes API.
* **Configuration flexible** : Pilotage complet depuis une feuille de calcul dédiée (aucun changement de code nécessaire pour modifier les paramètres).
* **Journalisation** : Historique des exécutions (durée, nombre d'erreurs) dans un onglet "Journal".
* **Rapports par email** : (Optionnel) Envoi d'un résumé d'exécution automatique.

## Prérequis et structure du Google Sheet

Pour que le script fonctionne, votre classeur Google Sheets doit contenir impérativement trois onglets.

### 1. Onglet `Configuration`
Cet onglet pilote le script. Créez les valeurs suivantes dans les colonnes A (Clé) et B (Valeur) :

| Colonne A (Paramètre) | Colonne B (Exemple de valeur) | Description |
| :--- | :--- | :--- |
| **Feuille de données** | `Audit` | Nom de l'onglet contenant les domaines à scanner. |
| **Colonne des domaines** | `A` | La lettre de la colonne où sont listés les domaines. |
| **Colonne de début des résultats** | `C` | La lettre de la colonne où le script commencera à écrire. |
| **Activer Verification MX** | `VRAI` | `VRAI` ou `FAUX`. |
| **Activer Verification SPF** | `VRAI` | `VRAI` ou `FAUX`. |
| **Activer Verification DMARC** | `VRAI` | `VRAI` ou `FAUX`. |
| **Activer Verification A** | `FAUX` | `VRAI` ou `FAUX`. |
| **Activer Rapport par Email** | `FAUX` | `VRAI` ou `FAUX`. |
| **Email pour le rapport** | `admin@example.com` | L'adresse de destination du rapport. |

### 2. Onglet de données (ex: `Audit`)
* Doit porter le nom défini dans la configuration.
* Doit contenir une liste de noms de domaines (ex: `google.com`) dans la colonne configurée, à partir de la ligne 2.
* La ligne 1 est réservée aux en-têtes (le script les générera automatiquement).

### 3. Onglet `Journal`
* Créez simplement un onglet vide nommé `Journal`. Le script y inscrira l'historique des exécutions.

## Installation

1.  Ouvrez votre Google Sheet.
2.  Allez dans **Extensions** > **Apps Script**.
3.  Copiez le contenu du fichier `Code.js` dans l'éditeur.
4.  Sauvegardez le projet (`Ctrl + S` ou `Cmd + S`).
5.  Rechargez votre Google Sheet (F5). Le menu **"Outils DNS Avancés"** apparaîtra après quelques secondes.

## Utilisation

### Mode manuel
1.  Cliquez sur le menu **Outils DNS Avancés** dans la barre d'outils.
2.  Sélectionnez **Lancer la vérification manuelle**.
3.  Lors de la première exécution, Google vous demandera d'autoriser le script (accès à UrlFetchApp et au Spreadsheet).

### Automatisation (Triggers)
Pour exécuter ce script périodiquement (ex: tous les matins) :
1.  Dans l'éditeur Apps Script, cliquez sur l'icône **Déclencheurs** (réveil) à gauche.
2.  Ajoutez un déclencheur :
    * Fonction : `logiqueTraitementDomaines` (ou `lancerVerificationManuelle`).
    * Source : "Déclenché par le temps".
    * Type : "Jour" ou "Heure".

## Notes techniques

* **API Rate Limits** : Le script utilise un délai de `200ms` (`DELAI_API_MS`) entre chaque appel pour respecter les quotas de l'API Google DNS publique.


