# Traitement des Bons ALDI

Application web pour découper un PDF multi-bons ALDI, détecter les signatures client et générer un récapitulatif Excel.

## Ce que ça fait

1. Tu déposes un PDF contenant des dizaines/centaines de bons d'intervention
2. L'app découpe chaque bon, détecte si le client a signé (tracé vectoriel PyMuPDF)
3. Tu télécharges un ZIP contenant :
   - `client_signe/` — PDFs des bons signés, nommés `BE00xxxx_code_dd-mm-yyyy.pdf`
   - `client_non_signe/` — PDFs des bons non signés
   - `recap_signes_non_signes.xlsx` — récapitulatif multi-onglets avec techniciens, filiales, doublons
   - `prestations_inconnues.txt` — libellés sans code reconnu (assignés `UNK`), à vérifier

## Option bonus : pointages

Si tu déposes aussi un fichier `Export Pointages.xlsx` et choisis le site ALDI, un onglet `Pointages_non_pointes` est ajouté à l'Excel.

## Stack

- **Frontend** : HTML/CSS/JS vanilla — drag-and-drop, aucune dépendance
- **Backend** : Netlify Function Python (AWS Lambda)
  - [PyMuPDF](https://pymupdf.readthedocs.io/) — découpe PDF + détection signature
  - [openpyxl](https://openpyxl.readthedocs.io/) — génération Excel
  - [pandas](https://pandas.pydata.org/) — traitement pointages

## Déploiement Netlify

```bash
netlify deploy --prod
```

Netlify détecte automatiquement les fonctions Python dans `netlify/functions/` et installe les dépendances via `requirements.txt`.

### Limites Netlify Functions (plan gratuit)
- **6 MB max** par requête (PDF + Excel)
- **10 s** de timeout (26 s sur les plans payants)

Pour des PDFs volumineux, migrer le backend sur Railway avec `dev_server.py` comme base FastAPI.

## Développement local

```bash
pip install fastapi uvicorn pymupdf openpyxl pandas python-multipart
python3 -m uvicorn dev_server:app --port 8765 --reload
# → http://localhost:8765
```
