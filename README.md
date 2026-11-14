# DocuVerify AI

Audit automatisé de documents d'import / export (facture, liste de colisage, connaissement, certificats…)
par un modèle d'IA multimodal. Interface en français et en anglais. Résultat : verdict, liste des 8 documents
obligatoires, validation croisée des données et actions à mener.

```
Navigateur ──► Vercel (frontend statique)
     │
     └────────► Render (backend FastAPI) ──► OpenRouter (modèle d'IA)
```

```
backend/    API FastAPI (main.py, requirements.txt, .env.example)
frontend/   Page web (index.html, config.js, static/app.css, vercel.json)
render.yaml Déploiement du backend sur Render (Blueprint)
```

## Lancer en local

```powershell
cd backend
pip install -r requirements.txt
$env:OPENROUTER_API_KEY = "votre-cle"
$env:OPENROUTER_MODEL = "nvidia/nemotron-3.5-super-vl:free"
python -m uvicorn main:app --port 8000
```

Ouvrez http://localhost:8000. En local, le backend sert aussi la page du dossier `frontend/`.

## Déploiement (Render + Vercel, offres gratuites)

**Avant de publier le dépôt sur GitHub :**
- ne commitez jamais `.env` ni une clé API (le `.gitignore` les exclut) ;
- ne commitez aucun document réel (PDF, Word, images) : ils contiennent des données d'entreprise ;
- vérifiez avec `git status` avant chaque `git commit`.

```bash
git init
git add .
git status          # pas de .env, pas de PDF
git commit -m "DocuVerify AI"
git branch -M main
git remote add origin https://github.com/VOTRE-COMPTE/docuverify.git
git push -u origin main
```

### 1. Backend sur Render

1. Render → **New → Blueprint**, choisissez votre dépôt. Render lit `render.yaml`.
2. Renseignez `OPENROUTER_API_KEY` (votre **nouvelle** clé) et, pour l'instant, `ALLOWED_ORIGINS` = `http://localhost:8000`.
3. Une fois déployé, notez l'adresse, par exemple `https://docuverify-api.onrender.com`.
   Contrôle : `https://…onrender.com/api/health` doit répondre `{"status":"ok",…}`.

### 2. Frontend sur Vercel

1. Dans `frontend/config.js`, mettez l'adresse Render :
   `window.DV_API_BASE = "https://docuverify-api.onrender.com";` puis commit et push.
2. Vercel → **Add New → Project**, importez le dépôt, **Root Directory = `frontend`**,
   Framework Preset **Other**, aucune commande de build.
3. Notez l'adresse Vercel, par exemple `https://docuverify.vercel.app`.

### 3. Autoriser le frontend

Dans Render → votre service → **Environment**, mettez
`ALLOWED_ORIGINS` = `https://docuverify.vercel.app` (votre adresse Vercel) et enregistrez (redéploiement automatique).
Pour accepter aussi les déploiements de prévisualisation Vercel, ajoutez
`ALLOWED_ORIGIN_REGEX` = `https://docuverify-.*\.vercel\.app`.

## Limites des offres gratuites

| Service | Limite | Effet |
|---|---|---|
| Render gratuit | 512 Mo de RAM, mise en veille après 15 min d'inactivité | premier accès ≈ 1 min ; la page affiche « Réveil du serveur… » |
| OpenRouter, modèles `:free` | 20 requêtes/min et **50 requêtes/jour** (1 000/jour après un achat unique de 10 $ de crédits) | un audit consomme 1 à 3 requêtes : comptez 15 à 25 audits par jour |
| OpenRouter | un solde négatif renvoie l'erreur 402, même pour les modèles gratuits | message d'erreur explicite dans l'interface |

Le plan gratuit de Render est réglé pour rester léger : 1 audit à la fois, 10 Mo par fichier, 20 pages
maximum, rendu à 1,5×. Les valeurs sont dans `render.yaml`.

## Passer au modèle payant

Dans Render → Environment : `OPENROUTER_MODEL` = `google/gemini-2.5-flash`, ajoutez des crédits sur OpenRouter,
et si vous voulez, remettez `MAX_ATTEMPTS` à `3`. Aucun changement de code.

## Sécurité : à savoir avant de rendre le dépôt public

- **Pas de connexion pour l'instant.** Quiconque connaît l'adresse du backend peut lancer des audits et consommer
  vos 50 requêtes quotidiennes. Une limite par adresse IP (10 audits/heure) et la restriction CORS réduisent le risque
  sans l'éliminer. Prochaine étape simple : un code d'accès partagé.
- **Modèles gratuits :** les fournisseurs peuvent enregistrer les requêtes pour entraîner leurs modèles. Pour les tests,
  utilisez des documents factices ; pour de vrais envois, passez au modèle payant et réglez la confidentialité
  dans OpenRouter.
- Les journaux du serveur ne contiennent pas le texte des documents (`LOG_CONTENT=0`).
- Le résultat de l'IA est un brouillon : un employé doit confirmer tampons, signatures et dates avant tout dépôt en douane.
- Choisissez une licence (par exemple MIT) si vous voulez autoriser la réutilisation du code.

## Variables d'environnement (backend)

| Variable | Rôle | Défaut |
|---|---|---|
| `OPENROUTER_API_KEY` | clé OpenRouter (obligatoire) | – |
| `OPENROUTER_MODEL` | modèle utilisé | `google/gemini-2.5-flash` |
| `ALLOWED_ORIGINS` | adresses autorisées à appeler l'API | localhost |
| `ALLOWED_ORIGIN_REGEX` | motif d'adresses autorisées (Vercel preview) | vide |
| `RATE_LIMIT_PER_HOUR` | audits par IP et par heure (0 = illimité) | 20 |
| `MAX_CONCURRENT_AUDITS` | audits simultanés | 2 |
| `MAX_FILE_MB`, `MAX_TOTAL_PAGES`, `PDF_ZOOM` | poids, pages et qualité d'image | 20, 40, 2.0 |
| `MAX_ATTEMPTS`, `MAX_OUTPUT_TOKENS` | tentatives et taille de réponse | 3, 8000 |
| `TRUST_PROXY` | 1 derrière Render/Vercel/proxy | 0 |
| `LOG_CONTENT` | 1 = extraits de documents dans les journaux (débogage) | 0 en production |

## Dépannage

| Symptôme | Cause probable |
|---|---|
| « Serveur injoignable » ou « Failed to fetch » | Render est en veille (attendez 1 min) ou `DV_API_BASE` est faux |
| Erreur CORS dans la console du navigateur | `ALLOWED_ORIGINS` ne contient pas l'adresse Vercel exacte (sans `/` final) |
| « Le modèle d'IA gratuit est limité… » | limite de 20/min ou de 50/jour atteinte : attendez ou passez au modèle payant |
| Erreur 402 | solde OpenRouter négatif |
| Le build Render échoue sur la version de Python | supprimez la variable `PYTHON_VERSION` dans `render.yaml` |
