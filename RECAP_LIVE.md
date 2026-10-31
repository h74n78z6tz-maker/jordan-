# Récap automatique des lives Whatnot

Ce sont les instructions que suit la tâche programmée. Elles servent aussi pour un récap à la main.

## Procédure

1. Lire `recaps/processed.json` : il contient les ID des lives déjà traités.
2. Avec le connecteur Whatnot, lister les lives au statut `ended` (`list_shows`, `status: ended`,
   `limit: 50`). Garder ceux qui ne sont pas dans `processed_show_ids`.
   S'il n'y en a aucun, s'arrêter sans rien committer.
3. Pour chaque nouveau live, du plus ancien au plus récent :
   1. Récupérer toutes les commandes (`list_orders`, `limit: 50`, en suivant `next_cursor`).
      Garder celles dont `show_id` est l'ID du live. Les commandes arrivent de la plus
      récente à la plus ancienne : on peut arrêter la pagination dès que `created_at` est
      antérieur au `scheduled_start_at` du live moins 1 jour.
   2. Pour chaque commande gardée, récupérer ses articles (`list_order_items`, `order_id`,
      `limit: 50`, en suivant `next_cursor`) et les mettre dans un champ `items` de la commande.
   3. Écrire le JSON brut dans `/tmp/recap_<show_id>.json` sous la forme
      `{"show": <objet show>, "orders": [<commande + items>, ...]}`.
      Ne pas modifier les valeurs renvoyées par l'API.
   4. Lancer `python3 scripts/build_recap.py /tmp/recap_<show_id>.json recaps/<AAAA-MM-JJ>_<titre-court>.xlsx`.
      La date est celle de `ended_at`. Le titre court reprend le titre du live, en minuscules,
      sans accents, avec des tirets, 40 caractères au maximum.
   5. Ajouter l'ID du live dans `recaps/processed.json`.
4. Faire un commit avec les nouveaux fichiers `.xlsx` et `processed.json`, message
   `Récap live : <titre>`, puis faire un push sur `master`.
5. Terminer par un court résumé en français pour chaque live : nombre de commandes, de clients
   et d'articles, total encaissé, et les 3 meilleurs clients.

## Règles

- L'API limite à environ 2 requêtes par seconde. En cas de réponse 429, attendre et réessayer.
- Lecture seule : ne jamais modifier, annuler ou créer quoi que ce soit sur Whatnot.
- Si un live terminé n'a aucune commande, générer quand même son récap (vide) et le marquer
  comme traité.

## Contenu du fichier Excel

- **Résumé** : titre du live, dates, nombre de commandes et de clients, articles vendus,
  chiffre d'affaires, frais de port, total et panier moyen.
- **Ventes** : une ligne par article vendu (date, commande, client, produit, SKU, type,
  quantité, prix, statut).
- **Clients** : total par client, trié du plus gros acheteur au plus petit.

Les commandes annulées ou échouées apparaissent dans « Ventes », mais ne comptent pas dans les totaux.
