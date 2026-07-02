# Redirection du domaine tournoibaseballrimouski.ca

## Objectif

Le domaine `tournoibaseballrimouski.ca` (et `www.tournoibaseballrimouski.ca`) ne
contient aucun site web. Il sert uniquement à rediriger automatiquement vers
l'affichage public du tournoi, publié comme application web Google Apps Script
(voir `doGet` / `showPublicUrl` dans `CLAUDE.md`).

## Architecture

```
Internet
   │
   ▼
DNS du fournisseur Internet
   │
   ▼
Cloudflare (DNS + SSL + Redirect)
   │
   ▼
Redirect Rule (301)
   │
   ▼
Google Apps Script
https://script.google.com/macros/s/AKfycbx2Zh7GqQGWhbCu_y0gDBoxS-4vg6XME_FiqksJe1mVWuB55XH9nx3Dnbk9LYFlZ9b1vQ/exec
```

## Webnames

Webnames sert **uniquement** à :

- Enregistrer le domaine
- Renouveler le domaine
- Configurer les nameservers

Les nameservers configurés chez Webnames sont :

```
cullen.ns.cloudflare.com
saanvi.ns.cloudflare.com
```

Une fois les nameservers configurés, **toute l'administration du domaine se
fait dans Cloudflare** — pas dans Webnames.

## Cloudflare

Cloudflare est responsable de :

- DNS
- Certificat SSL
- Redirection HTTP
- Proxy
- Sécurité
- Cache

Toute modification doit être faite dans Cloudflare, jamais dans Webnames.

### DNS actuel

| Type | Nom | Cible | Proxy |
|------|-----|-------|-------|
| A | `@` | `209.15.37.6` | ON (nuage orange) |
| CNAME | `www` | `tournoibaseballrimouski.ca` | ON |
| TXT | — | `v=spf1 -all` | — |

Le serveur `209.15.37.6` est celui de Webnames. Ce n'est pas un problème
puisque Cloudflare intercepte les requêtes **avant** qu'elles n'atteignent ce
serveur (proxy activé).

### Redirect Rule

Une *Redirect Rule* Cloudflare effectue une redirection permanente (301) vers
Google Apps Script.

**Condition** : `Host = tournoibaseballrimouski.ca` OU
`Host = www.tournoibaseballrimouski.ca`

**Action** : 301 Permanent Redirect

**Destination** :
`https://script.google.com/macros/s/AKfycbx2Zh7GqQGWhbCu_y0gDBoxS-4vg6XME_FiqksJe1mVWuB55XH9nx3Dnbk9LYFlZ9b1vQ/exec`

La règle doit être :

- **Enabled**
- **Placée en première position** (Place at: First)

> ⚠️ **Important** — Si la Redirect Rule est désactivée, Cloudflare laisse
> passer la requête vers le serveur Webnames et on obtient la page
> « This website is under development ». C'est exactement le problème
> rencontré le 2 juillet 2026.

### Workers

Les *Workers* ne sont plus nécessaires. Les anciennes Worker Routes suivantes
peuvent être supprimées si elles existent encore :

```
tournoibaseballrimouski.ca/*
www.tournoibaseballrimouski.ca/*
```

## Google Apps Script

Déploiement Web App (voir aussi la section « Public display web app » de
`CLAUDE.md`) :

- **Execute as** : Me (`danieltremblay18@gmail.com`)
- **Who has access** : Anyone
- **URL** : `https://script.google.com/macros/s/AKfycbx2Zh7GqQGWhbCu_y0gDBoxS-4vg6XME_FiqksJe1mVWuB55XH9nx3Dnbk9LYFlZ9b1vQ/exec`

> Rappel : chaque changement au code de l'affichage public nécessite un
> **nouveau déploiement** (Déployer › Nouveau déploiement) — coller le code
> seul ne met pas à jour l'URL `/exec` déjà en ligne.

## Vérifications

Vérifier que les nameservers sont toujours :

```
cullen.ns.cloudflare.com
saanvi.ns.cloudflare.com
```

Vérifier que le DNS passe par Cloudflare :

```
nslookup tournoibaseballrimouski.ca
```

Résultat attendu — une adresse appartenant à Cloudflare : `172.64.x.x`,
`104.x.x.x`, ou `2606:4700:...`

### Ping

Si `ping tournoibaseballrimouski.ca` ne répond pas, c'est normal : Cloudflare
ne répond généralement pas aux requêtes ICMP (ping).

## Si la redirection ne fonctionne plus

1. Vérifier les nameservers chez Webnames.
2. Vérifier que la Redirect Rule est **ENABLED**.
3. Vérifier qu'elle est en **première position**.
4. Vérifier que les DNS sont toujours en mode **Proxy** (nuage orange).
5. Vérifier que le déploiement Google Apps Script est toujours actif.
6. Tester directement l'URL Google Apps Script.

## À retenir

- **Webnames** = propriétaire du domaine (registrar).
- **Cloudflare** = DNS, SSL, proxy et redirection.
- **Google Apps Script** = application web (affichage public).

Le navigateur ne communique jamais directement avec Webnames. Il passe
toujours par Cloudflare, qui redirige ensuite vers Google Apps Script.
