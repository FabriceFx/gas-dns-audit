/**
 * @OnlyCurrentDoc
 *
 * Script avancé de vérification DNS, piloté par une feuille de configuration.
 * Fonctionnalités :
 * - Configuration flexible (feuilles, colonnes, types de vérification).
 * - Interprétation des enregistrements DMARC et SPF.
 * - Journal d'exécution automatique.
 * - Envoi de rapports par e-mail.
 * - Validation de la configuration avec messages d'erreur clairs.
 * - Mécanisme anti-cache fiable pour des données fraîches.
 * - Écriture robuste dans le journal.
 */

// --- CONSTANTES GLOBALES ---
const CONFIG_SHEET_NAME = 'Configuration';
const LOG_SHEET_NAME = 'Journal';
const MESSAGES_STATUT = {
  ERREUR_API: "Erreur API",
  DOMAINE_INVALIDE: "Domaine invalide",
  PAS_ENREGISTREMENT_MX: "Pas de MX",
  PAS_ENREGISTREMENT_SPF: "Pas de SPF",
  PAS_ENREGISTREMENT_DMARC: "Pas de DMARC",
  PAS_ENREGISTREMENT_A: "Pas de A",
  FOURNISSEUR_EMAIL_GOOGLE: "Google Workspace",
  FOURNISSEUR_EMAIL_OFFICE365: "Office 365",
  AUTRE_FOURNISSEUR_EMAIL: "Autre",
};
const DELAI_API_MS = 200;

// --- GESTION DE L'INTERFACE (MENU) ---

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Outils DNS Avancés')
    .addItem('Lancer la vérification manuelle', 'lancerVerificationManuelle')
    .addToUi();
}

/**
 * Fonction appelée par le menu. Gère l'UI et appelle la logique principale.
 * Inclut un bloc try...catch pour afficher les erreurs de configuration.
 */
function lancerVerificationManuelle() {
  const ui = SpreadsheetApp.getUi();
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Démarrage du traitement...', 'Initialisation', 10);
    
    const resume = logiqueTraitementDomaines();
    
    if (!resume) return; // Arrêt si la logique a échoué au démarrage

    if (resume.domainesTraites > 0) {
      ui.alert('Traitement terminé', `${resume.domainesTraites} domaine(s) traité(s) avec succès.\n${resume.erreurs} erreur(s) rencontrée(s).`, ui.ButtonSet.OK);
    } else {
      ui.alert('Aucun domaine traité', 'Vérifiez la configuration ou le contenu de la colonne des domaines.', ui.ButtonSet.OK);
    }
  } catch (e) {
    // Affiche les erreurs de configuration de manière claire à l'utilisateur
    ui.alert('Erreur Critique', e.message, ui.ButtonSet.OK);
  }
}

// --- LOGIQUE PRINCIPALE (POUR DÉCLENCHEUR ET MANUEL) ---

/**
 * Cœur du script. Lit la config, traite les domaines, écrit les résultats et les journaux.
 * @return {object|null} Un objet résumé de l'exécution, ou null si la configuration est invalide.
 */
function logiqueTraitementDomaines() {
  let config;
  try {
    config = recupererConfiguration();
    Logger.log("Configuration lue : " + JSON.stringify(config, null, 2));
  } catch (e) {
    Logger.log(`ERREUR DE CONFIGURATION: ${e.message}`);
    // Propage l'erreur pour que l'exécution manuelle puisse l'afficher.
    throw e;
  }
  
  const debutExecution = new Date();
  const feuilleDonnees = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.nomFeuilleDonnees);
  if (!feuilleDonnees) {
    throw new Error(`La feuille de données "${config.nomFeuilleDonnees}" (définie dans la configuration) est introuvable.`);
  }

  const enTetes = ['Domaine'];
  if (config.activerVerificationMX) enTetes.push('Fournisseur MX');
  if (config.activerVerificationSPF) enTetes.push('SPF', 'SPF Policy');
  if (config.activerVerificationDMARC) enTetes.push('DMARC', 'DMARC Policy');
  if (config.activerVerificationA) enTetes.push('Adresse A');

  feuilleDonnees.getRange(1, config.colonneDebutResultats, 1, enTetes.length).setValues([enTetes]).setFontWeight('bold');
  //feuilleDonnees.getRange(2, config.colonneDebutResultats - 1).setValue(`Dernière vérification:\n${debutExecution.toLocaleString('fr-FR')}`);

  const derniereLigne = feuilleDonnees.getLastRow();
  if (derniereLigne < 2) {
    Logger.log('Aucun domaine à traiter.');
    return { domainesTraites: 0, erreurs: 0 };
  }
  
  const plageDomaines = feuilleDonnees.getRange(2, config.colonneDomaines, derniereLigne - 1, 1);
  const valeursDomaines = plageDomaines.getValues();

  const resultats = [];
  let nombreDomainesTraites = 0;
  let nombreErreurs = 0;

  for (const [domaine] of valeursDomaines) {
    const ligneResultat = [domaine || ''];
    if (domaine && typeof domaine === 'string' && domaine.trim() !== "") {
      const domaineNettoye = domaine.trim();
      
      try {
        if (config.activerVerificationMX) {
          ligneResultat.push(rechercherMX(domaineNettoye));
        }
        if (config.activerVerificationSPF) {
          const {enregistrement, politique} = interpreterSPF(domaineNettoye);
          ligneResultat.push(enregistrement, politique);
        }
        if (config.activerVerificationDMARC) {
          const {enregistrement, politique} = interpreterDMARC(domaineNettoye);
          ligneResultat.push(enregistrement, politique);
        }
        if (config.activerVerificationA) {
          ligneResultat.push(rechercherA(domaineNettoye));
        }

        nombreDomainesTraites++;
      } catch (e) {
        Logger.log(`Erreur irrécupérable pour le domaine ${domaineNettoye}: ${e}`);
        nombreErreurs++;
        // En cas d'erreur, on remplit le reste de la ligne pour garder la cohérence
        while(ligneResultat.length < enTetes.length) ligneResultat.push("Erreur Script");
      }
    }

    // CORRECTION : S'assurer que chaque ligne a le bon nombre de colonnes, même les lignes vides.
    while (ligneResultat.length < enTetes.length) {
      ligneResultat.push(''); // Ajouter des chaînes vides pour combler
    }
    
    resultats.push(ligneResultat);
  }
  
  if (resultats.length > 0) {
    feuilleDonnees.getRange(2, config.colonneDebutResultats, resultats.length, resultats[0].length).setValues(resultats);
  }

  const finExecution = new Date();
  const dureeMs = finExecution - debutExecution;
  const resume = {
      date: finExecution.toLocaleString('fr-FR'),
      domainesTraites: nombreDomainesTraites,
      erreurs: nombreErreurs,
      duree: `${(dureeMs / 1000).toFixed(2)}s`,
  };

  ecrireJournal(resume);
  
  Logger.log(`Vérification de l'envoi d'e-mail. Activation: ${config.activerRapportParEmail}, Email: ${config.emailPourLeRapport}`);

  if (config.activerRapportParEmail && config.emailPourLeRapport) {
    envoyerRapportParEmail(resume, config.emailPourLeRapport);
  }

  return resume;
}

// --- FONCTIONS DE RECHERCHE ET INTERPRÉTATION ---

function rechercherMX(domaine) {
  const reponses = recupererDonneesDNS(domaine, "MX");
  if (reponses === null) return MESSAGES_STATUT.ERREUR_API;
  if (reponses.length === 0) return MESSAGES_STATUT.PAS_ENREGISTREMENT_MX;
  const enregistrementsMX = reponses.map(rep => rep.data.toLowerCase()).join(' ');
  if (enregistrementsMX.includes("google.com") || enregistrementsMX.includes("googlemail.com")) return MESSAGES_STATUT.FOURNISSEUR_EMAIL_GOOGLE;
  if (enregistrementsMX.includes("outlook.com")) return MESSAGES_STATUT.FOURNISSEUR_EMAIL_OFFICE365;
  return MESSAGES_STATUT.AUTRE_FOURNISSEUR_EMAIL;
}

function interpreterSPF(domaine) {
  const reponses = recupererDonneesDNS(domaine, "TXT");
  if (reponses === null) return { enregistrement: MESSAGES_STATUT.ERREUR_API, politique: '' };
  const enregistrementSPF = reponses.find(r => r.data && r.data.toLowerCase().startsWith("v=spf1"));
  if (!enregistrementSPF) return { enregistrement: MESSAGES_STATUT.PAS_ENREGISTREMENT_SPF, politique: '' };
  const enregistrementNettoye = enregistrementSPF.data.replace(/"/g, '');
  let politique = 'Non spécifiée';
  if (enregistrementNettoye.includes('-all')) politique = 'Hard Fail (-all)';
  else if (enregistrementNettoye.includes('~all')) politique = 'Soft Fail (~all)';
  else if (enregistrementNettoye.includes('?all')) politique = 'Neutral (?all)';
  return { enregistrement: enregistrementNettoye, politique: politique };
}

function interpreterDMARC(domaine) {
  const domaineDMARC = `_dmarc.${domaine}`;
  const reponses = recupererDonneesDNS(domaineDMARC, "TXT");
  if (reponses === null) return { enregistrement: MESSAGES_STATUT.ERREUR_API, politique: '' };
  const enregistrementDMARC = reponses.find(r => r.data && r.data.toLowerCase().startsWith("v=dmarc1"));
  if (!enregistrementDMARC) return { enregistrement: MESSAGES_STATUT.PAS_ENREGISTREMENT_DMARC, politique: '' };
  const enregistrementNettoye = enregistrementDMARC.data.replace(/"/g, '');
  const matchPolitique = enregistrementNettoye.match(/p=([^;]+)/);
  const politique = matchPolitique ? matchPolitique[1].trim().toLowerCase() : 'Aucune';
  return { enregistrement: enregistrementNettoye, politique: `p=${politique}` };
}

function rechercherA(domaine) {
  const reponses = recupererDonneesDNS(domaine, "A");
  if (reponses === null) return MESSAGES_STATUT.ERREUR_API;
  if (reponses.length === 0) return MESSAGES_STATUT.PAS_ENREGISTREMENT_A;
  return reponses[0].data || MESSAGES_STATUT.PAS_ENREGISTREMENT_A;
}

// --- OUTILS ET UTILITAIRES ---

function recupererDonneesDNS(nom, type) {
  const randomPadding = Math.random().toString(36).substring(2);
  const url = `https://dns.google/resolve?name=${encodeURIComponent(nom)}&type=${type}&random_padding=${randomPadding}`;
  
  try {
    Utilities.sleep(DELAI_API_MS);
    const resultat = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resultat.getResponseCode() !== 200) {
      Logger.log(`Erreur HTTP ${resultat.getResponseCode()} pour ${nom} [${type}]`);
      return null;
    }
    const reponse = JSON.parse(resultat.getContentText());
    if (reponse.Status !== 0 && reponse.Status !== 3) { // 3 = NXDOMAIN
      Logger.log(`Erreur de statut DNS ${reponse.Status} pour ${nom} [${type}]`);
      return null;
    }
    return reponse.Answer || [];
  } catch (e) {
    Logger.log(`Exception lors de la récupération DNS pour ${nom} [${type}]: ${e}`);
    return null;
  }
}

/**
 * Lit la configuration et la valide. Lève une erreur si la configuration est invalide.
 * @return {object} L'objet de configuration validé.
 */
function recupererConfiguration() {
  const feuilleConfig = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  if (!feuilleConfig) {
    throw new Error(`Feuille de configuration introuvable. Veuillez créer une feuille nommée exactement "${CONFIG_SHEET_NAME}".`);
  }
  const data = feuilleConfig.getRange("A2:B10").getValues();
  const config = {};
  data.forEach(([cle, valeur]) => {
    if (cle) {
      const cleCamelCase = cle.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s(.)/g, (match, p1) => p1.toUpperCase()).replace(/\s/g, '').replace(/^(.)/, (match, p1) => p1.toLowerCase());
      // Convertir les valeurs textuelles "VRAI" ou "Oui" en booléens
      if (typeof valeur === 'string' && valeur.toLowerCase() === 'vrai' || valeur.toLowerCase() === 'oui') {
          config[cleCamelCase] = true;
      } else if (typeof valeur === 'string' && valeur.toLowerCase() === 'faux' || valeur.toLowerCase() === 'non') {
          config[cleCamelCase] = false;
      } else {
        config[cleCamelCase] = valeur;
      }
    }
  });

  config.nomFeuilleDonnees = config.feuilleDeDonnees;

  if (!config.nomFeuilleDonnees) throw new Error('Configuration manquante : La valeur pour "Feuille de données" est manquante ou mal orthographiée dans l\'onglet Configuration.');
  if (!config.colonneDesDomaines) throw new Error('Configuration manquante : La valeur pour "Colonne des domaines" est manquante ou mal orthographiée.');
  if (!config.colonneDeDebutDesResultats) throw new Error('Configuration manquante : La valeur pour "Colonne de début des résultats" est manquante ou mal orthographiée.');
  if (!config.hasOwnProperty('activerRapportParEmail')) throw new Error('Configuration manquante : La clé "Activer Rapport par Email" est manquante ou mal orthographiée.');
  if (!config.hasOwnProperty('emailPourLeRapport')) throw new Error('Configuration manquante : La clé "Email pour le rapport" est manquante ou mal orthographiée.');

  try {
    config.colonneDomaines = feuilleConfig.getRange(config.colonneDesDomaines + "1").getColumn();
    config.colonneDebutResultats = feuilleConfig.getRange(config.colonneDeDebutDesResultats + "1").getColumn();
  } catch(e) {
    throw new Error(`Erreur de configuration : La valeur pour une des colonnes ("${config.colonneDesDomaines}" ou "${config.colonneDeDebutDesResultats}") est invalide. Veuillez utiliser une lettre de colonne valide (ex: A, B, C).`);
  }
  
  return config;
}

function ecrireJournal(resume) {
  const feuilleJournal = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (feuilleJournal) {
    const premiereLigneVide = feuilleJournal.getLastRow() + 1;
    const ligneJournal = [resume.date, resume.domainesTraites, resume.erreurs, resume.duree];
    
    if (feuilleJournal.getLastRow() === 0) {
        feuilleJournal.getRange("A1:D1").setValues([["Date", "Domaines traités", "Erreurs", "Durée d'exécution"]]).setFontWeight('bold');
    }
    feuilleJournal.getRange(premiereLigneVide, 1, 1, ligneJournal.length).setValues([ligneJournal]);
  }
}

function envoyerRapportParEmail(resume, email) {
  const sujet = `Rapport de vérification DNS - ${new Date().toLocaleDateString('fr-FR')}`;
  const corps = `
    <h2>Rapport d'exécution du script de vérification DNS</h2>
    <p>Le script a terminé son exécution le ${resume.date}.</p>
    <ul>
      <li><b>Domaines traités :</b> ${resume.domainesTraites}</li>
      <li><b>Erreurs rencontrées :</b> ${resume.erreurs}</li>
      <li><b>Durée du traitement :</b> ${resume.duree}</li>
    </ul>
    <p>Vous pouvez consulter les résultats détaillés dans votre feuille de calcul.</p>
  `;
  MailApp.sendEmail({ to: email, subject: sujet, htmlBody: corps });
}
