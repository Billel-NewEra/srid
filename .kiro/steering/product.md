---
inclusion: always
---

# Produit — SRID COM

## Vue d'ensemble
**SRID COM** est une application web de gestion **financière** et **logistique** pour les
sociétés **SRID** et **SRID Genetics** (société algérienne de produits phytosanitaires /
agrochimiques).

- **Développeur** : MobiBenz (https://mobibenz.dz/fr)
- **Utilisateurs** : équipe administrative (admin, saisie) + consultants en lecture seule.
- **Hébergement** : cPanel (Passenger WSGI), déployé sous un sous-chemin (ex. `mondomaine.com/app`).

## Objectifs métier
- Fournir un **dashboard** clair pour les décisions de gestion.
- Garder la visibilité **SRID** et **SRID Genetics** explicite partout.
- Suivre les **opérations financières** : chèques, virements, versements, transferts.
- Suivre la **chaîne logistique** : commandes fournisseurs, import (dates douane D10,
  arrivée, échéances de paiement).
- Maintenir une **haute qualité de données** (normalisation des banques, filtres fiables).
- UX simple, prévisible et visuellement cohérente.

## Domaines fonctionnels
1. **Finance** — saisie, consultation, filtres, statuts de chèques, échéances, rejets.
2. **Logistique** — bons de commande, gestion des commandes (D10 → arrivée → échéance →
   paiement), frais et prix de revient.
3. **Référentiels** — clients, remettants, fournisseurs (par société), produits (par société).
4. **Dashboard** — KPI + graphiques (mensuel, par société, par statut, top 5 clients).
5. **Administration** — utilisateurs, rôles, journal d'audit (historique).

## Rôles
| Rôle | Droits |
|------|--------|
| `admin` | Tout (CRUD, users, suppression, statuts) |
| `saisie` | Création / modification + référentiels ; pas de suppression ni gestion users |
| `consultation` | Lecture seule |

## Règles produit importantes
- Les deux sociétés doivent rester distinctes : valeur société canonique = **`SRID`** ou
  **`Genetics`** dans la table `operations` (attention aux variantes de casse à l'import).
- Ne pas retirer de comportement de filtre existant sans demande explicite.
- Filtres déterministes et prévisibles pour l'utilisateur.
- Préférer des changements UI ciblés et minimes pour éviter les régressions.
