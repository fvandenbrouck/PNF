# Polices auto-hébergées

Les polices utilisées par les widgets sont servies depuis ce dossier, et non plus depuis Google Fonts, afin qu'aucune adresse IP d'utilisateur ne soit transmise à un tiers à l'affichage.

| Police | Usage | Fichiers | Source |
|---|---|---|---|
| Fraunces | titres | `fraunces-var.woff2`, `fraunces-italic-400.woff2` | https://github.com/undercasetype/Fraunces |
| DM Sans | texte courant | `dm-sans-var.woff2` | https://github.com/googlefonts/dm-fonts |
| DM Mono | étiquettes | `dm-mono-400.woff2`, `dm-mono-500.woff2` | https://github.com/googlefonts/dm-fonts |

Les trois polices sont distribuées sous **SIL Open Font License 1.1** (https://openfontlicense.org), qui autorise leur redistribution avec les applications. Seul le sous-ensemble **latin** est embarqué ; les caractères rares hors de ce sous-ensemble s'affichent avec la police système de repli.

Pour mettre à jour : télécharger les fichiers `woff2` correspondants depuis la feuille de style Google Fonts (sous-ensemble `latin`) et conserver les noms de fichiers.
