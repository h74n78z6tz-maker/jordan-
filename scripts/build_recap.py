#!/usr/bin/env python3
"""Génère le récap Excel d'un live Whatnot.

Entrée : un fichier JSON avec les réponses brutes de l'API vendeur Whatnot :
  {
    "show":   { ...ExternalShow... },
    "orders": [ { ...ExternalOrder..., "items": [ ...ExternalOrderItem... ] } ]
  }
Sortie : un fichier .xlsx avec 3 onglets (Ventes, Clients, Résumé).

Usage : python3 scripts/build_recap.py donnees.json recaps/2026-11-02_mon-live.xlsx
"""
import json
import sys
from collections import defaultdict
from datetime import datetime
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

STATUTS = {
    "pending": "En attente",
    "pending_fulfillment": "À expédier",
    "completed": "Terminée",
    "cancelled": "Annulée",
    "failed": "Échouée",
}
TYPES = {
    "buy_it_now": "Achat immédiat",
    "auction": "Enchère",
    "giveaway": "Giveaway",
    "other": "Autre",
}
ENTETE = Font(bold=True, color="FFFFFF")
FOND = PatternFill("solid", fgColor="1F4E78")
EUROS = '#,##0.00 "€"'


def money(m):
    if not m or m.get("amount") in (None, ""):
        return Decimal("0")
    return Decimal(str(m["amount"]))


def date_fr(iso):
    if not iso:
        return ""
    try:
        return datetime.fromisoformat(iso.replace("Z", "+00:00")).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return iso


def client(order):
    b = order.get("buyer") or {}
    return b.get("username") or "(inconnu)", b.get("display_name") or ""


def feuille(ws, entetes, lignes, colonnes_euros=(), largeurs=None):
    ws.append(entetes)
    for c in ws[1]:
        c.font, c.fill = ENTETE, FOND
        c.alignment = Alignment(horizontal="center", vertical="center")
    for ligne in lignes:
        ws.append(ligne)
    for idx in colonnes_euros:
        for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
            for c in row:
                c.number_format = EUROS
    for i, h in enumerate(entetes, start=1):
        largeur = (largeurs or {}).get(h)
        if largeur is None:
            vals = [len(str(h))] + [len(str(r[i - 1])) for r in lignes if r[i - 1] is not None]
            largeur = min(max(vals) + 2, 60)
        ws.column_dimensions[get_column_letter(i)].width = largeur
    ws.freeze_panes = "A2"
    if lignes:
        ws.auto_filter.ref = ws.dimensions


def build(data, sortie):
    show = data.get("show") or {}
    orders = [o for o in data.get("orders", []) if o.get("show_id") in (None, show.get("id"))]

    ventes, par_client = [], defaultdict(lambda: {"nom": "", "cmd": set(), "articles": 0,
                                                  "sous_total": Decimal(0), "port": Decimal(0),
                                                  "total": Decimal(0)})
    tot = {"articles": 0, "sous_total": Decimal(0), "port": Decimal(0), "taxes": Decimal(0),
           "total": Decimal(0), "annulees": 0, "commandes": 0}

    for o in sorted(orders, key=lambda o: o.get("created_at") or ""):
        pseudo, nom = client(o)
        statut = o.get("status")
        annulee = statut in ("cancelled", "failed")
        a = o.get("amounts") or {}
        items = o.get("items") or []
        for it in items:
            pi = it.get("purchased_item") or {}
            qte = int(it.get("quantity") or 1)
            pu = money(it.get("unit_price"))
            ventes.append([
                date_fr(o.get("created_at")), o.get("id"), pseudo, nom,
                pi.get("title") or it.get("product_id") or "", pi.get("sku") or "",
                TYPES.get(o.get("transaction_type"), o.get("transaction_type") or ""),
                qte, float(pu), float(pu * qte), STATUTS.get(statut, statut or ""),
            ])
        if not items:
            ventes.append([date_fr(o.get("created_at")), o.get("id"), pseudo, nom, "(détail indisponible)",
                           "", TYPES.get(o.get("transaction_type"), ""), None, None,
                           float(money(a.get("subtotal"))), STATUTS.get(statut, statut or "")])
        if annulee:
            tot["annulees"] += 1
            continue
        n_art = sum(int(it.get("quantity") or 1) for it in items)
        c = par_client[pseudo]
        c["nom"] = nom
        c["cmd"].add(o.get("id"))
        c["articles"] += n_art
        c["sous_total"] += money(a.get("subtotal"))
        c["port"] += money(a.get("shipping"))
        c["total"] += money(a.get("total"))
        tot["commandes"] += 1
        tot["articles"] += n_art
        for k, champ in (("sous_total", "subtotal"), ("port", "shipping"), ("taxes", "tax"), ("total", "total")):
            tot[k] += money(a.get(champ))

    wb = Workbook()
    ws = wb.active
    ws.title = "Ventes"
    feuille(ws, ["Date", "N° commande", "Client (pseudo)", "Nom", "Produit", "SKU", "Type",
                 "Qté", "Prix unitaire", "Total ligne", "Statut"], ventes, colonnes_euros=(9, 10))

    lignes_clients = sorted(
        ([p, c["nom"], len(c["cmd"]), c["articles"], float(c["sous_total"]), float(c["port"]), float(c["total"])]
         for p, c in par_client.items()), key=lambda r: -r[6])
    feuille(wb.create_sheet("Clients"), ["Client (pseudo)", "Nom", "Commandes", "Articles",
                                         "Sous-total", "Port", "Total payé"], lignes_clients,
            colonnes_euros=(5, 6, 7))

    ws = wb.create_sheet("Résumé")
    resume = [
        ("Live", show.get("title") or ""),
        ("Lien", show.get("url") or ""),
        ("Début prévu", date_fr(show.get("scheduled_start_at"))),
        ("Fin", date_fr(show.get("ended_at"))),
        ("Commandes valides", tot["commandes"]),
        ("Commandes annulées / échouées", tot["annulees"]),
        ("Clients", len(par_client)),
        ("Articles vendus", tot["articles"]),
        ("Chiffre d'affaires (hors port)", float(tot["sous_total"])),
        ("Frais de port", float(tot["port"])),
        ("Taxes", float(tot["taxes"])),
        ("Total encaissé", float(tot["total"])),
        ("Panier moyen", float(tot["total"] / tot["commandes"]) if tot["commandes"] else 0.0),
    ]
    for k, v in resume:
        ws.append([k, v])
        ws.cell(ws.max_row, 1).font = Font(bold=True)
        if isinstance(v, float):
            ws.cell(ws.max_row, 2).number_format = EUROS
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 60
    wb.move_sheet("Résumé", offset=-2)
    wb.active = 0
    wb.save(sortie)
    return tot, len(par_client)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    with open(sys.argv[1], encoding="utf-8") as f:
        donnees = json.load(f)
    t, n = build(donnees, sys.argv[2])
    print(f"OK {sys.argv[2]} : {t['commandes']} commandes, {n} clients, "
          f"{t['articles']} articles, total {t['total']} €")
