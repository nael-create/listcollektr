# listcollektr

Listes de prospects pour Collektr.

## Asie du Sud-Est : vinyles & cartes Pokémon (octobre 2026)

- `Leads_Asie_SE_Vinyles_Pokemon.xlsx` : le fichier à utiliser.
  - **Résumé** : compteurs par pays (formules), légende et méthode.
  - **Vinyles** : disquaires et vendeurs de vinyles (Thaïlande, Vietnam, Philippines, Singapour, Indonésie).
  - **Pokémon** : boutiques et revendeurs de cartes Pokémon, avec taille, preuves de « gros items » (booster boxes, cases, slabs PSA…) et priorité A/B/C.
- `data/` : les mêmes données en CSV (import CRM) et en JSON.
- `scripts/build_xlsx.py` : fusionne les fichiers de recherche JSON, dédoublonne et régénère l'Excel.

```bash
python scripts/build_xlsx.py --research <dossier_json> --out Leads_Asie_SE_Vinyles_Pokemon.xlsx --data-dir data
```

Chaque ligne cite ses sources. Les contacts viennent de recherches web publiques : vérifier le compte avant prise de contact.

## Lots musique : 250 lots entre RM45 et RM60 (octobre 2026)

- `Lots_Musique_250_RM45-60.xlsx` : 88 vinyles (35 %), 62 CD d'albums cultes (25 %), 75 Funko Pop musique, 13 livres illustrés, 12 autres objets. Funko Pop jusqu'à RM70 (25 achetables en Malaisie, le reste à l'étranger). Chaque ligne a son prix en RM, son lien et la preuve de stock (relevée le 20/10/2026). Onglet « Réserve » : remplaçants en cas de rupture.
- `data/lots_musique.csv` / `.json` : mêmes données ; `data/research_lots_musique/` : relevés bruts par catégorie.

```bash
python scripts/build_lots_musique.py --research data/research_lots_musique --out Lots_Musique_250_RM45-60.xlsx --data-dir data
```
