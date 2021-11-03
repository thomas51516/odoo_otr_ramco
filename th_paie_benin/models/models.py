# -*- coding: utf-8 -*-

from odoo import models, api, _, fields
from odoo.exceptions import ValidationError
from math import *
from datetime import *


class hrContactInherit(models.Model):
    _inherit = 'hr.contract'

    heure_sup_12_jour = fields.Float(
        string='HEURES SUPP JOUR 12%',
        default=0,
    )
    heure_sup_35_jour = fields.Float(
        string='HEURES SUPP JOUR 35%',
        default=0,
    )
    heure_sup_50_jour = fields.Float(
        string='HEURES SUPP JOUR 50%',
        default=0,
    )
    heure_sup_50_jour_ferier = fields.Float(
        string='HEURES SUPP JOUR FERIER 50%',
        default=0,
    )
    heure_sup_100_nuit = fields.Float(
        string='HEURES SUPP NUIT 100%',
        default=0,
    )
    prime_rendement = fields.Float(
        string='Prime de rendement',
        default=0,
    )
    provision_cp = fields.Float(
        string='Provision CP',
        default=0,
    )
    appel_urgence = fields.Float(
        string="Appel d'urgence",
        default=0,
    )
    astreinte = fields.Float(
        string='Astreinte',
        default=0,
    )

    indemnite_transport = fields.Float(
        string='Indemnité de transport',
        default=0,
    )
    prime_caisse = fields.Float(
        string='Prime de caisse',
        default=0,
    )
    prime_garde = fields.Float(
        string='Prime de garde',
        default=0,
    )
    prime_panier = fields.Float(
        string='Prime de panier',
        default=0,
    )
    prime_risque = fields.Float(
        string='Prime de rique',
    )
    prime_salisure = fields.Float(
        string="Prime de salissure"
    )
    prime_speciale = fields.Float(
        string='Prime spéciale',
    )
    prime_specialite = fields.Float(
        string='Prime de spécialité',
    )
    prime_responsabilite = fields.Float(
        string='Prime de responsabilté',
    )
    rapel_salaire_imp = fields.Float(
        string='Rappel de salaire imposable',
    )
    remboursement_pret = fields.Float(
        string='Remboursement de prêt',
    )
    sursalaire = fields.Float(
        string='Sursalaire',
    )
    regularisation_salaire_net = fields.Float(
        string="Régularisation salaire net",
    )
    prime_sujetion = fields.Float(
        string="Prime de sujétion"
    )
    indemnite_deloignement = fields.Float(
        string="Indemnité d’éloignement"
    )
    prime_de_performance = fields.Float(
        string="Prime de performance"
    )

    nombre_heure = fields.Float(
        string="Nombre d'heure"
    )

    taux_horaire = fields.Float(
        string="Taux horaire"
    )

    type_paiement = fields.Selection([
        ('espece', 'Espèce'),
        ('virement', 'Virement'),
        ('cheque', 'Chèque banciare'),
    ], string="Type de paiement", default="virement")


class hrPersoneACharge(models.Model):
    _name = 'hr.personne.acharge'
    _description = 'Personne à charge'

    name = fields.Char(
        string='Nom et prénom(s)',
        required=True,
    )
    date_naissance = fields.Date(string="Date de naissance")
    age = fields.Integer(
        string='Age',
        store=True,
    )

    employee_id = fields.Many2one(
        'hr.employee',
        string='Personne à charge',
    )

    @api.onchange('date_naissance')
    def _calculer_age(self):
        for rec in self:
            if type(rec.date_naissance) != bool:
                today_year = datetime.today().year
                birth_year = rec.date_naissance.year
                age = today_year - birth_year
                rec.age = age


class hrEmployeInherit(models.Model):
    _inherit = 'hr.employee'

    persone_acharge_ids = fields.One2many(
        'hr.personne.acharge',
        'employee_id',
        string='Personnes à charge',
    )
    numero_cnss = fields.Char(
        string="Numéro de CNSS",
    )
    matricule = fields.Char(
        string='Numéro matricule',
    )
    payslip_ids = fields.One2many(
        'hr.payslip',
        'employee_id',
        string='Field Label',
    )
    heure_travail = fields.Float(
        string="Nombre d'heure de travail",
    )
    count_bancaire = fields.Char(
        string="Numéro de compte"
    )
    banque = fields.Char(
        string="Banque"
    )
    indice = fields.Float(
        string="Indice"
    )
    valeur_indice = fields.Float(
        string="Valeur indice"
    )

    @api.onchange('persone_acharge_ids')
    def onchange_persone_acharge_ids(self):
        for rec in self:
            rec.children = len(rec.persone_acharge_ids)


class RubriqueLine(models.Model):
    _name = 'rubrique.lier'
    _description = 'Description'

    nom = fields.Char(
        string='Rubrique',
    )
    montant = fields.Char(
        string='Montant',
    )
