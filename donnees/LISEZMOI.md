# Jeux de données

Dossier des jeux de données du dépôt `Meca-IA_Ed2026` (Méca-IA).

**Avertissement :** les jeux de données de ce dépôt ont un sens mécanique donné par transformation : les noms de colonnes, et parfois les valeurs, sont modifiés pour obtenir des ordres de grandeur réalistes. Certains partent de jeux publics connus, d'autres sont générés par l'auteur. Les métamodèles et les lois issus des estimateurs d'apprentissage automatique de ces notebooks ne doivent pas servir en application réelle. Pour refaire l'apprentissage, chargez vos propres données : c'est la méthode qui compte.

## Format

Un jeu est un classeur Excel `MecaIA-<identifiant>.xlsx` : une table, une ligne d'en-têtes (noms de colonnes), une ligne par observation. Les notebooks de ce dépôt ne lisent que les jeux de ce dossier, sans jamais le modifier ; un fichier de ce dossier ne se retouche pas, il se copie.

```
import pandas as pd
df = pd.read_excel("donnees/MecaIA-<identifiant>.xlsx")
```

## Sur Colab

Un notebook ouvert seul dans Colab ne voit pas ce dossier. Sa première cellule récupère dans le dépôt public les jeux qu'il lit ; sinon, ouvrez le panneau des fichiers de Colab, créez un dossier `donnees` à la racine de l'espace de travail et téléversez-y les classeurs utiles. La cellule « Préparation du dépôt (Colab) » de chaque notebook vérifie leur présence et liste ce qui manque. Voir la section « Dans Colab » du [`README`](../README.md).

## Où trouver le reste

- Le tableau des jeux, avec leur cible, leurs dimensions et leur origine publique : section « Jeux de données » du [`README`](../README.md).
- Les empreintes sha256 et les tailles : [`manifest.json`](../manifest.json).
- Les origines, les licences d'origine et les modifications de chaque jeu : [`LICENSE-DONNEES.md`](../LICENSE-DONNEES.md) (licence CC BY 4.0, sous réserve de la licence d'origine de chaque jeu).
