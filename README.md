# Calendrier escrime consolidé

Une page HTML statique qui regroupe, pour la saison en cours, le circuit national FFE, les épreuves de la ligue Nouvelle-Aquitaine, les épreuves de Gironde et les stages. On peut filtrer par arme, catégorie, sexe et type d’épreuve, et afficher un mois en calendrier, la saison en liste, ou toute l’année sur une seule page (vue « Année », un mois par colonne). Le jour courant est entouré en rouge.

La page est régénérée **chaque lundi** par GitHub Actions et publiée sur GitHub Pages.

## Comment ça marche

```
site de la ligue ──► scripts/build_calendar.py ──► dist/index.html ──► GitHub Pages
  (Excel + agenda)         ▲          ▲
                           │          └── template/calendar.html (mise en page, filtres, vues)
                           └── cdescrime33.fr (Gironde) + data/evenements_manuels.json (corrections)
```

Le script `scripts/build_calendar.py` :

1. lit la page <https://www.escrime-nouvelle-aquitaine.fr/events/> et télécharge les fichiers Excel qui y sont liés. Il reconnaît tout seul la **grille FFE** (circuit national, vacances scolaires) et le **calendrier des épreuves Nouvelle-Aquitaine** (épreuves régionales et stages), même si leur nom change en cours de saison ;
2. lit l’**agenda du site de la ligue** (API de l’extension « The Events Calendar ») et ajoute les épreuves qui ne figurent pas dans les Excel ;
3. lit le **calendrier du comité de Gironde** (<https://cdescrime33.fr/calendrier/>) et les épreuves saisies à la main dans `data/evenements_manuels.json` ;
4. fusionne le tout, détecte les doublons et les dates qui divergent entre sources (affichées « À confirmer »), puis écrit `dist/index.html` et `dist/evenements.json`.

Chaque source téléchargée est conservée dans `data/sources/`. Si le site de la ligue est en panne, la dernière copie est utilisée : la page n’est jamais vidée. La saison est détectée automatiquement (à partir d’août, on passe à la saison suivante) ; on peut la forcer avec `"saison": "2026-2027"` dans `config.json`.

Ne sont pas repris : escrime artistique, para-escrime et circuits internationaux (FIE, EFC).

## Épreuves de Gironde

Les rencontres départementales sont lues automatiquement sur la page du comité, <https://cdescrime33.fr/calendrier/> (réglage `gironde_page` dans `config.json`). Cette page est un simple texte mis en forme : le script repère les titres de section (« Rencontres départementales d’épée », etc.), puis chaque date et le lieu écrit à côté ou juste en dessous. « 23 ou 30 janvier » est placé au 23 avec la mention « À confirmer ». Les rencontres sont ouvertes à toutes les catégories, de M9 à vétérans. Si le comité change la mise en page et qu’aucune date n’est plus reconnue, les points à vérifier de la page le signalent.

## Ajouter ou corriger une épreuve à la main

Pour une épreuve absente de toutes les sources, éditer `data/evenements_manuels.json` (directement sur GitHub, bouton ✏️) :

```json
{
  "titre": "Circuit départemental 1",
  "lieu": "Pessac",
  "debut": "2026-12-06",
  "fin": "2026-12-06",
  "niveau": "gironde",
  "armes": ["sabre", "epee"],
  "categories": ["M11", "M13"],
  "sexe": "",
  "detail": "Individuel",
  "a_confirmer": null,
  "lien": null
}
```

- `niveau` : `gironde`, `region`, `france` ou `stage` ;
- `armes` : `epee`, `sabre`, `fleuret`, `animateur` ;
- `sexe` : `dames`, `hommes` ou vide ;
- `a_confirmer` : un texte affiché en bandeau jaune, ou `null`.

La section `masquer` permet de retirer un événement généré à tort, par exemple `{"debut": "2027-01-23", "titre_contient": "Pau"}`.

Chaque modification sur `main` relance la génération ; la page est à jour une minute plus tard.

## Lancer en local

```bash
pip install -r requirements.txt
python scripts/build_calendar.py            # télécharge les sources
python scripts/build_calendar.py --offline  # n’utilise que data/sources/
# puis ouvrir dist/index.html
```

## Mise en ligne sur GitHub Pages

1. Fusionner la PR dans `main`.
2. Dans **Settings › Pages**, choisir **Source : GitHub Actions**.
3. Lancer une première fois **Actions › Calendrier escrime › Run workflow** (ensuite c’est automatique chaque lundi, et à chaque modification sur `main`).

La page est alors visible sur <https://fsmjla.github.io/fencingCalendar/>.

## Brancher la page sur pessac-escrime.fr

Aujourd’hui, `www.pessac-escrime.fr` ne sert pas de site : c’est une **redirection** (HTTP 302) vers `https://www.spuc-escrime.com/`, le site du club hébergé chez SportsRegions. Or GitHub Pages ne sait servir un domaine que **en entier** : on ne peut pas lui confier seulement le chemin `/calendar` d’un domaine géré ailleurs. Trois options :

### Option A, recommandée pour obtenir exactement `www.pessac-escrime.fr/calendar`

On confie `www.pessac-escrime.fr` à GitHub Pages, qui continue de rediriger vers le site du club, et le calendrier vient se loger dans `/calendar`.

1. **Créer un dépôt `fsmjla.github.io`** (le « site utilisateur ») contenant :
   - `index.html` qui redirige vers le site du club :
     ```html
     <!doctype html><meta charset="utf-8">
     <meta http-equiv="refresh" content="0; url=https://www.spuc-escrime.com/">
     <link rel="canonical" href="https://www.spuc-escrime.com/">
     <a href="https://www.spuc-escrime.com/">SPUC Escrime</a>
     ```
   - le même contenu dans `404.html`, pour que toute autre adresse continue de rediriger vers le club.
2. Dans ce dépôt, **Settings › Pages** : source « Deploy from a branch » (`main`, `/`), puis **Custom domain : `www.pessac-escrime.fr`**.
3. **Renommer ce dépôt-ci en `calendar`** (Settings › General › Repository name). Un site de projet hérite du domaine du site utilisateur : la page sera servie sur `https://www.pessac-escrime.fr/calendar/`, sans autre réglage.
4. **DNS**, chez le registraire du domaine :
   - supprimer la redirection web actuelle de `www` (et de l’apex si elle existe) ;
   - `www` : enregistrement **CNAME** vers `fsmjla.github.io.` ;
   - `pessac-escrime.fr` (apex) : quatre enregistrements **A** vers `185.199.108.153`, `185.199.109.153`, `185.199.110.153`, `185.199.111.153` (GitHub redirige alors l’apex vers `www`).
5. Une fois le DNS propagé (quelques minutes à quelques heures), cocher **Enforce HTTPS** dans Settings › Pages du dépôt `fsmjla.github.io`.
6. Conseillé : **vérifier le domaine** dans les paramètres du compte GitHub (Settings › Pages › Add a verified domain), pour qu’aucun autre compte ne puisse l’utiliser.

### Option B, la plus simple : un sous-domaine

Adresse finale du type `https://calendrier.pessac-escrime.fr/`, sans toucher à la redirection actuelle.

1. DNS : `calendrier` en **CNAME** vers `fsmjla.github.io.`
2. Dans ce dépôt, **Settings › Pages › Custom domain** : `calendrier.pessac-escrime.fr`, puis **Enforce HTTPS**.

### Option C, en complément : l’intégrer au site du club

Sur le site SportsRegions, la page « Calendrier de la saison » peut afficher le calendrier dans un cadre, si l’éditeur accepte un bloc HTML ou « iframe » :

```html
<iframe src="https://www.pessac-escrime.fr/calendar/" style="width:100%;height:1400px;border:0" title="Calendrier escrime"></iframe>
```

À défaut, un simple lien vers la page fait l’affaire.

## Fichiers

| Chemin | Rôle |
| --- | --- |
| `scripts/build_calendar.py` | lecture des sources, fusion, génération de la page |
| `template/calendar.html` | mise en page, filtres, vues mois, liste et année, version compacte sur smartphone (données injectées à la place de `__DATA__`) |
| `config.json` | adresses des sources, saison forcée éventuelle |
| `data/evenements_manuels.json` | ajouts et corrections manuels |
| `data/sources/` | dernières copies des sources (secours si le site de la ligue est indisponible) |
| `.github/workflows/calendrier.yml` | génération hebdomadaire et publication sur GitHub Pages |
