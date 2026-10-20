# Tri des fiches de paie

Petit programme **100 % local** qui lit des fiches de paie en PDF (numériques, pas des scans), les range par année / mois / employeur et calcule le **net imposable par année**, pour aider à remplir une déclaration de revenus en retard ou oubliée.

> Rien n'est envoyé sur Internet. Les PDF d'origine ne sont jamais modifiés : le programme en fait des copies rangées dans un nouveau dossier.

## Ce que ça fait

- Lit les fiches de paie de plusieurs mises en page différentes (agences d'intérim, petites entreprises, etc.).
- Copie chaque fiche dans `2024/`, `2025/`… sous la forme `2025-10_Employeur.pdf`.
- Produit une **fiche de report** (page à ouvrir ou imprimer) : pour chaque année, les montants à recopier sur la déclaration de revenus (formulaire 2042), avec un indicateur de fiabilité :
  - **1AJ** : traitements et salaires (net imposable) — **1BJ** pour le conjoint, **1CJ** / **1DJ** pour les personnes à charge ;
  - **8HV** : prélèvement à la source déjà payé (colonne « Impôt prélevé » des bulletins) — 8IV, 8JV, 8KV selon la personne ;
  - **1GH** : heures supplémentaires exonérées (colonne « HS/HC exonér »), seulement s'il y en a — 1HH, 1IH, 1JH selon la personne.
  Ces numéros sont ceux du formulaire des revenus 2023 : vérifiez-les sur le formulaire de chaque année.
- Produit un tableau Excel avec 5 onglets :
  - **Aide déclaration** : la même chose que la fiche de report, avec le détail par employeur ;
  - **Récap annuel** : net imposable total par année et par employeur, comparé au cumul indiqué sur la dernière fiche quand il est lisible ;
  - **Détail** : une ligne par fiche (brut, net imposable, net à payer…) ;
  - **Mois manquants** : les mois sans fiche entre la première et la dernière fiche d'un employeur ;
  - **À vérifier** : fiches mal lues, scans, doublons.

## Ce que ça ne fait pas

- Il **ne remplit pas** la déclaration (papier ou en ligne) et ne se connecte pas à impots.gouv.fr : il prépare la fiche de report, c'est à vous de recopier les montants.
- Il ne calcule pas les frais réels (case 1AK), les allocations chômage (1AP) ni les pensions (1AS), et ne tient pas compte des exonérations particulières (apprentis, étudiants…).
- Il **ne lit pas les scans ou photos** (pas d'OCR) : ils sont mis de côté dans « À vérifier ».
- Ce n'est **pas un conseil fiscal**. Vérifiez toujours les montants : comparez avec le cumul de la dernière fiche de l'année et avec les revenus déjà pré-remplis sur votre déclaration ou votre espace impots.gouv.fr. Pour une situation en retard, le service des impôts (SIP) peut vous indiquer la marche à suivre.

## Utilisation (Windows, le plus simple)

1. Dans l'onglet **Releases** de ce dépôt GitHub, téléchargez `TriFichesPaie.exe`.
2. Double-cliquez dessus. Windows peut afficher un avertissement « SmartScreen » car le programme n'est pas signé : cliquez sur *Informations complémentaires* puis *Exécuter quand même*. Le code source est visible ici si vous voulez vérifier.
3. Cliquez sur **Parcourir…**, choisissez le dossier qui contient les fiches de paie (les sous-dossiers sont lus aussi), indiquez de qui sont ces fiches (déclarant 1, conjoint…), puis **Lancer le tri**.
4. À la fin, cliquez sur **Ouvrir la fiche de report** : c'est la page à avoir sous les yeux en remplissant la déclaration. Le tableau Excel donne le détail.
5. Si une année est marquée **À vérifier**, regardez pourquoi (mois sans fiche, écart avec le cumul…) avant de recopier le montant.

À chaque lancement, un nouveau dossier `Fiches_de_paie_triees_<date>` est créé : rien n'est écrasé.

## Utilisation avec Python

```bash
pip install -r requirements.txt
python app.py                                   # version avec fenêtre
python trier_fiches_paie.py "C:\mes\fiches" -o resultat   # en ligne de commande
python trier_fiches_paie.py "C:\mes\fiches" -o resultat --declarant 2   # fiches du conjoint (case 1BJ)
```

## Si une fiche est mal lue (mise en page inconnue)

1. Lancez avec l'option diagnostic (case à cocher dans la fenêtre, ou `--debug` en ligne de commande) : le texte lu dans chaque PDF est enregistré dans un dossier `debug/`.
2. Repérez comment le *net imposable*, la *période* ou l'*employeur* y sont écrits.
3. Créez un fichier `patterns.json` **à côté du programme** :

```json
{
  "net_imposable": ["montant fiscal"],
  "employeur": ["agence\\s*:\\s*(.+)"]
}
```

Les clés possibles : `net_imposable`, `cumul_net_imposable`, `brut`, `net_a_payer`, `employeur`. Ce sont des expressions régulières, ajoutées en priorité aux motifs existants.

## Développement

```bash
pip install -r requirements-dev.txt
pytest
```

Le fichier `.exe` est construit automatiquement par GitHub Actions quand on publie une version (étiquette `v1.0.0`, etc.).

## Données personnelles

Les fiches de paie contiennent des données sensibles. Le `.gitignore` de ce dépôt exclut les PDF, les tableaux Excel et les dossiers de résultat pour éviter de les publier par erreur. Ne les ajoutez jamais au dépôt, ne les joignez pas à une *issue*, et masquez nom et numéro de sécurité sociale avant de partager un exemple.

## Licence

MIT, voir `LICENSE`.
