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
