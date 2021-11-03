
import base64
import os
from datetime import datetime
from datetime import *
from io import BytesIO

import xlsxwriter
from PIL import Image as Image
from odoo import fields, models, api, _
from odoo.exceptions import ValidationError
from xlsxwriter.utility import xl_rowcol_to_cell
import re


class ExcelBilan(models.TransientModel):
    _name = 'bilan.excel.wiz'
    date_fin = fields.Date(
        string="A la date du",
        default=fields.date.today(),
    )
    est_comptabilise = fields.Boolean(
        string="Inclure les écritures non comptabilisées"
    )

    def get_bilan_data(self):
        data = {}
        state = 'posted'
        if self.est_comptabilise == True:
            state = 'draft'
        liste_ecriture_comptable = []
        ecriture_comptable = self.env['account.move.line'].search(
            [('date', '<=', self.date_fin)])
        for e in ecriture_comptable:
            if e.move_id.state == state or e.move_id.state == 'posted':
                vals = {
                    'account_id': e.account_id.code,
                    'credit': e.credit,
                    'debit': e.debit,
                    'balance': e.balance,
                }
                liste_ecriture_comptable.append(vals)

        # Actif du bilan
        # immobilisations incorporelles
        immobilisation_incorporel = 0
        immobilisation_incorporel_amortissement = 0

        frais_de_developpement_et_de_prospection = 0
        frais_de_developpement_et_de_prospection_amortissement = 0

        brevets_licence_logiciel_et_droits_similaires = 0
        brevets_licence_logiciel_et_droits_similaires_amortissement = 0

        fonds_commercial_et_droit_au_bail = 0
        fonds_commercial_et_droit_au_bail_amortissement = 0

        autre_immobilisation_corporelles = 0
        autre_immobilisation_corporelles_amortissement = 0

        # immobilisations corporelles
        immobilisation_corporel = 0
        immobilisation_corporel_amortissement = 0

        terrains = 0
        terrains_amortissement = 0

        batiment = 0
        batiment_amortissement = 0

        amenagements_agencements_installations = 0
        amenagements_agencements_installations_amortissement = 0

        materiel_mobilier_et_actifs_biologiques = 0
        materiel_mobilier_et_actifs_biologiques_amortissement = 0

        materiel_de_transport = 0
        materiel_de_transport_amortissement = 0

        avance_et_acomptes_verse_sur_immobilisation = 0
        avance_et_acomptes_verse_sur_immobilisation_amortissement = 0

        # immobilisation financière
        immobilisation_financiere = 0
        immobilisation_financiere_amortissement = 0

        titre_de_participation = 0
        titre_de_participation_amortissement = 0

        autres_immobilisation_financiere = 0
        autres_immobilisation_financiere_amortissement = 0

        # actif ciculant
        actif_circulant_hao = 0
        actif_circulant_hao_amortissement = 0

        stock_et_en_cours = 0
        stock_et_en_cours_amortissement = 0
        #
        creances_et_emplois_assimiles = 0
        creances_et_emplois_assimiles_amortissement = 0

        fournisseurs_avances_versees = 0
        fournisseurs_avances_versees_amortissement = 0

        clients = 0
        clients_amortissement = 0

        autres_creances = 0
        autres_creances_amortissement = 0

        total_actif_cirulant = 0
        total_actif_cirulant_amortissement = 0

        # tresorerie actif
        titres_de_placement = 0
        titres_de_placement_amortissement = 0

        valeur_a_encaisser = 0
        valeur_a_encaisser_amortissement = 0

        banque_cheque_postaux_caisse_et_assimile = 0
        banque_cheque_postaux_caisse_et_assimile_amortissement = 0

        total_tresorerie_actif = 0
        total_tresorerie_actif_amortissement = 0

        # ecart_de_conversion_actif
        ecart_de_conversion_actif = 0

        # PASSIF DU BILAN
        # CAPRO RA
        capital = 0
        apporteurs_captial_non_appele = 0
        prime_liees_au_capital_social = 0
        ecarts_de_reevaluation = 0
        reserve_indisponible = 0
        reserve_libre = 0
        report_a_nouveau = 0
        resultat_net_de_exercice = 0
        subvention_investissement = 0
        provision_reglementees = 0
        total_capitaux_propres_et_ressources_assimilees = 0

        # DEFI RA
        emprunts_et_dettes_financieres_divers = 0
        dettes_de_location_acquisition = 0
        provisions_pour_risques_et_charges = 0
        total_dettes_financieres_et_ressources_assimilees = 0

        # Total resource stable
        total_ressources_stables = 0

        # PASSIF CIRCULANT
        dettes_circulantes_hao = 0
        clients_avance_recues = 0
        fournisseur_dexploitation = 0
        dettes_fiscales_et_sociales = 0
        autres_dettes = 0
        provision_pour_risque_a_court_terme = 0

        # Totla passif circulant
        total_passif_circulant = 0

        # tresorerie passif
        banque_credit_d_escompte = 0
        banque_etablissement_financier_et_credit_de_tresorerie = 0

        # total tresorerie passif
        total_tresorerie_passif = 0

        # ecart de converion passif
        ecart_de_converion_passif = 0

        # total general passif
        total_general_passif = 0

        for ecriture in liste_ecriture_comptable:
            # immobilisation incorporel
            # frais de developpement et de prospection 211, 2181, 2191
            if re.match('^211', str(ecriture['account_id'])) or re.match('^2181', str(ecriture['account_id'])) or re.match('^2191', str(ecriture['account_id'])):
                frais_de_developpement_et_de_prospection += ecriture['balance']

             # amortissement
            if re.match('2811', str(ecriture['account_id'])) or re.match('2818', str(ecriture['account_id'])) or re.match('2911', str(ecriture['account_id'])) or re.match('2918', str(ecriture['account_id'])) or re.match('2919', str(ecriture['account_id'])):
                frais_de_developpement_et_de_prospection_amortissement += ecriture['balance']

            if re.match('^212', str(ecriture['account_id'])) or re.match('^213', str(ecriture['account_id'])) or re.match('^214', str(ecriture['account_id'])) or re.match('^2193', str(ecriture['account_id'])):
                brevets_licence_logiciel_et_droits_similaires += ecriture['balance']

            # amortissment
            if re.match('2812', str(ecriture['account_id'])) or re.match('2813', str(ecriture['account_id'])) or re.match('2814', str(ecriture['account_id'])) or re.match('2912', str(ecriture['account_id'])) or re.match('2913', str(ecriture['account_id'])) or re.match('2914', str(ecriture['account_id'])) or re.match('2919', str(ecriture['account_id'])):
                brevets_licence_logiciel_et_droits_similaires_amortissement += ecriture['balance']

            if re.match('^215', str(ecriture['account_id'])) or re.match('^216', str(ecriture['account_id'])):
                fonds_commercial_et_droit_au_bail += ecriture['balance']
            # amortissement
            if re.match('2815', str(ecriture['account_id'])) or re.match('2816', str(ecriture['account_id'])) or re.match('2915', str(ecriture['account_id'])) or re.match('2916', str(ecriture['account_id'])):
                fonds_commercial_et_droit_au_bail_amortissement += ecriture['balance']

            if re.match('^217', str(ecriture['account_id'])) or re.match('^218', str(ecriture['account_id'])) and not re.match('^2181', str(ecriture['account_id'])) or re.match('^2198', str(ecriture['account_id'])):
                autre_immobilisation_corporelles += ecriture['balance']
            # amortissement
            if re.match('2817', str(ecriture['account_id'])) or re.match('2818', str(ecriture['account_id'])) or re.match('2917', str(ecriture['account_id'])) or re.match('2918', str(ecriture['account_id'])) or re.match('2919', str(ecriture['account_id'])):
                autre_immobilisation_corporelles_amortissement += ecriture['balance']

            # immobilisation corporel
            if re.match('^22', str(ecriture['account_id'])):
                terrains += ecriture['balance']

            # amortissement
            if re.match('^282', str(ecriture['account_id'])) or re.match('^292', str(ecriture['account_id'])):
                terrains_amortissement += ecriture['balance']

            if re.match('^231', str(ecriture['account_id'])) or re.match('^232', str(ecriture['account_id'])) or re.match('^233', str(ecriture['account_id'])) or re.match('^237', str(ecriture['account_id'])) or re.match('^2391', str(ecriture['account_id'])):
                batiment += ecriture['balance']
            # amortissement
            if re.match('^2831', str(ecriture['account_id'])) or re.match('^2832', str(ecriture['account_id'])) or re.match('^2833', str(ecriture['account_id'])) or re.match('^2837', str(ecriture['account_id'])) or re.match('^2931', str(ecriture['account_id'])) or re.match('^2932', str(ecriture['account_id'])) or re.match('^2933', str(ecriture['account_id'])) or re.match('^2937', str(ecriture['account_id'])) or re.match('^2939', str(ecriture['account_id'])):
                batiment_amortissement += ecriture['balance']

            if re.match('^234', str(ecriture['account_id'])) or re.match('^235', str(ecriture['account_id'])) or re.match('^238', str(ecriture['account_id'])) or re.match('^2392', str(ecriture['account_id'])) or re.match('^2393', str(ecriture['account_id'])):
                amenagements_agencements_installations += ecriture['balance']

            # amortissement
            if re.match('^2834', str(ecriture['account_id'])) or re.match('^2835', str(ecriture['account_id'])) or re.match('^2838', str(ecriture['account_id'])) or re.match('^2934', str(ecriture['account_id'])) or re.match('^2935', str(ecriture['account_id'])) or re.match('^2938', str(ecriture['account_id'])) or re.match('^2939', str(ecriture['account_id'])):
                amenagements_agencements_installations_amortissement += ecriture['balance']

            if re.match('^24', str(ecriture['account_id'])) and not re.match('^245', str(ecriture['account_id'])) or re.match('^2495', str(ecriture['account_id'])):
                materiel_mobilier_et_actifs_biologiques += ecriture['balance']
            # amortissement
            if re.match('^284', str(ecriture['account_id'])) and not re.match('^2845', str(ecriture['account_id'])) or re.match('^294', str(ecriture['account_id'])) and not re.match('^2945', str(ecriture['account_id'])) and not re.match('^2949', str(ecriture['account_id'])) or re.match('^2949', str(ecriture['account_id'])):
                materiel_mobilier_et_actifs_biologiques_amortissement += ecriture['balance']

            if re.match('^245', str(ecriture['account_id'])) or re.match('^2495', str(ecriture['account_id'])):
                materiel_de_transport += ecriture['balance']
            # amortissement
            if re.match('2845', str(ecriture['account_id'])) or re.match('2945', str(ecriture['account_id'])) or re.match('2949', str(ecriture['account_id'])):
                materiel_de_transport_amortissement += ecriture['balance']

            if re.match('^251', str(ecriture['account_id'])) or re.match('^252', str(ecriture['account_id'])):
                avance_et_acomptes_verse_sur_immobilisation += ecriture['balance']
            # amortissement
            if re.match('^2951', str(ecriture['account_id'])) or re.match('^2952', str(ecriture['account_id'])):
                avance_et_acomptes_verse_sur_immobilisation_amortissement += ecriture['balance']

            # immobilisation financière
            if re.match('^26', str(ecriture['account_id'])):
                titre_de_participation += ecriture['balance']

            # amortissement
            if re.match('^296', str(ecriture['account_id'])):
                titre_de_participation_amortissement += ecriture['balance']

            if re.match('^27', str(ecriture['account_id'])):
                autres_immobilisation_financiere += ecriture['balance']
            # amortissement
            if re.match('^297', str(ecriture['account_id'])):
                autres_immobilisation_financiere_amortissement += ecriture['balance']

            # actif ciculant
            if re.match('^485', str(ecriture['account_id'])) or re.match('^488', str(ecriture['account_id'])):
                actif_circulant_hao += ecriture['balance']

            # amortissement
            if re.match('^498', str(ecriture['account_id'])):
                actif_circulant_hao_amortissement += ecriture['balance']

            if re.match('^31', str(ecriture['account_id'])) or re.match('^32', str(ecriture['account_id'])) or re.match('^33', str(ecriture['account_id'])) or re.match('^34', str(ecriture['account_id'])) or re.match('^35', str(ecriture['account_id'])) or re.match('^36', str(ecriture['account_id'])) or re.match('^37', str(ecriture['account_id'])) or re.match('^38', str(ecriture['account_id'])):
                stock_et_en_cours += ecriture['balance']

            # amortissement
            if re.match('^39', str(ecriture['account_id'])):
                stock_et_en_cours_amortissement += ecriture['balance']

            if re.match('^409', str(ecriture['account_id'])):
                fournisseurs_avances_versees += ecriture['balance']

            # amortissement
            if re.match('^490', str(ecriture['account_id'])):
                fournisseurs_avances_versees_amortissement += ecriture['balance']

            if re.match('^41', str(ecriture['account_id'])) and not re.match('^419', str(ecriture['account_id'])):
                clients += ecriture['balance']

            # amortissement
            if re.match('^491', str(ecriture['account_id'])):
                clients_amortissement += ecriture['balance']

            if re.match('^185', str(ecriture['account_id'])) and not re.match('^478', str(ecriture['account_id'])) or re.match('^42', str(ecriture['account_id'])) or re.match('^43', str(ecriture['account_id'])) or re.match('^44', str(ecriture['account_id'])) or re.match('^45', str(ecriture['account_id'])) or re.match('^46', str(ecriture['account_id'])) or re.match('^47', str(ecriture['account_id'])):
                if ecriture['balance'] > 0:
                    autres_creances += ecriture['balance']

            # amortissement
            if re.match('^492', str(ecriture['account_id'])) or re.match('^493', str(ecriture['account_id'])) or re.match('^494', str(ecriture['account_id'])) or re.match('^495', str(ecriture['account_id'])) or re.match('^496', str(ecriture['account_id'])) or re.match('^497', str(ecriture['account_id'])):
                autres_creances_amortissement += ecriture['balance']

             # Total
            creances_et_emplois_assimiles = fournisseurs_avances_versees + \
                clients + autres_creances

            # amortissement
            creances_et_emplois_assimiles_amortissement = fournisseurs_avances_versees_amortissement + \
                clients_amortissement + autres_creances_amortissement

            # tresorerie actif
            if re.match('^50', str(ecriture['account_id'])):
                titres_de_placement += ecriture['balance']

            # amortissement
            if re.match('^590', str(ecriture['account_id'])):
                titres_de_placement_amortissement += ecriture['balance']

            if re.match('^51', str(ecriture['account_id'])):
                valeur_a_encaisser += ecriture['balance']
            # amortissement
            if re.match('^591', str(ecriture['account_id'])):
                valeur_a_encaisser_amortissement += ecriture['balance']

            if re.match('^52', str(ecriture['account_id'])) or re.match('^57', str(ecriture['account_id'])) or re.match('^53', str(ecriture['account_id'])) or re.match('^54', str(ecriture['account_id'])) or re.match('^55', str(ecriture['account_id'])) or re.match('^581', str(ecriture['account_id'])) or re.match('^582', str(ecriture['account_id'])):
                if ecriture['balance'] > 0:
                    banque_cheque_postaux_caisse_et_assimile += ecriture['balance']
            # amortissement
            if re.match('^592', str(ecriture['account_id'])) or re.match('^593', str(ecriture['account_id'])) or re.match('^594', str(ecriture['account_id'])):
                banque_cheque_postaux_caisse_et_assimile_amortissement += ecriture['balance']

            # total_tresorerie_actif
            total_tresorerie_actif = titres_de_placement + \
                valeur_a_encaisser + banque_cheque_postaux_caisse_et_assimile
            # amortissement
            total_tresorerie_actif_amortissement = titres_de_placement_amortissement + \
                valeur_a_encaisser_amortissement + \
                banque_cheque_postaux_caisse_et_assimile_amortissement

            if re.match('^478', str(ecriture['account_id'])):
                ecart_de_conversion_actif += ecriture['balance']

            # PASSIF DU BILAN
            # CAPRO RA
            if re.match('^104', str(ecriture['account_id'])) or re.match('^101', str(ecriture['account_id'])):
                capital += ecriture['balance']

            if re.match('^109', str(ecriture['account_id'])):
                apporteurs_captial_non_appele += ecriture['balance']

            if re.match('^105', str(ecriture['account_id'])):
                prime_liees_au_capital_social += ecriture['balance']

            if re.match('^106', str(ecriture['account_id'])):
                ecarts_de_reevaluation += ecriture['balance']

            if re.match('^111', str(ecriture['account_id'])) or re.match('^112', str(ecriture['account_id'])) or re.match('^113', str(ecriture['account_id'])):
                reserve_indisponible += ecriture['balance']

            if re.match('^118', str(ecriture['account_id'])):
                reserve_libre += ecriture['balance']

            if re.match('^12', str(ecriture['account_id'])):
                report_a_nouveau += ecriture['balance']

            if re.match('^13', str(ecriture['account_id'])):
                resultat_net_de_exercice += ecriture['balance']

            if re.match('^14', str(ecriture['account_id'])):
                subvention_investissement += ecriture['balance']

            if re.match('^15', str(ecriture['account_id'])):
                provision_reglementees += ecriture['balance']

            # DEFI RAA
            if re.match('^16', str(ecriture['account_id'])) or re.match('^181', str(ecriture['account_id'])) or re.match('^182', str(ecriture['account_id'])) or re.match('^183', str(ecriture['account_id'])) or re.match('^184', str(ecriture['account_id'])):
                emprunts_et_dettes_financieres_divers += ecriture['balance']

            if re.match('^17', str(ecriture['account_id'])):
                dettes_de_location_acquisition += ecriture['balance']

            if re.match('^19', str(ecriture['account_id'])):
                provisions_pour_risques_et_charges += ecriture['balance']

            # PASSIF CIRCULANT

            if re.match('481', str(ecriture['account_id'])) or re.match('482', str(ecriture['account_id'])) or re.match('484', str(ecriture['account_id'])) or re.match('4998', str(ecriture['account_id'])):
                dettes_circulantes_hao += ecriture['balance']

            if re.match('^419', str(ecriture['account_id'])):
                clients_avance_recues += ecriture['balance']

            if re.match('^40', str(ecriture['account_id'])) and not re.match('^409', str(ecriture['account_id'])):
                fournisseur_dexploitation += ecriture['balance']

            if re.match('^42', str(ecriture['account_id'])) or re.match('^43', str(ecriture['account_id'])) or re.match('^44', str(ecriture['account_id'])):
                if ecriture['balance'] < 0:
                    dettes_fiscales_et_sociales += ecriture['balance']

            if re.match('^185', str(ecriture['account_id'])) or re.match('^599', str(ecriture['account_id'])) or re.match('^45', str(ecriture['account_id'])) or re.match('^47', str(ecriture['account_id'])) and not re.match('^479', str(ecriture['account_id'])):
                if ecriture['balance'] < 0:
                    autres_dettes += ecriture['balance']

            if re.match('^499', str(ecriture['account_id'])) or re.match('^45', str(ecriture['account_id'])) and not re.match('^4998', str(ecriture['account_id'])):
                provision_pour_risque_a_court_terme += ecriture['balance']

            if re.match('^564', str(ecriture['account_id'])) or re.match('^565', str(ecriture['account_id'])):
                banque_credit_d_escompte += ecriture['balance']

            if re.match('^52', str(ecriture['account_id'])) or re.match('^53', str(ecriture['account_id'])) or re.match('^561', str(ecriture['account_id'])) or re.match('^566', str(ecriture['account_id'])):
                if ecriture['balance'] < 0:
                    banque_etablissement_financier_et_credit_de_tresorerie += ecriture['balance']

            if re.match('^479', str(ecriture['account_id'])):
                ecart_de_converion_passif += ecriture['balance']

        # Totaux
        # immobilisation incorporel
        immobilisation_incorporel = frais_de_developpement_et_de_prospection + \
            brevets_licence_logiciel_et_droits_similaires + \
            fonds_commercial_et_droit_au_bail + autre_immobilisation_corporelles

        # amortissement
        immobilisation_incorporel_amortissement = frais_de_developpement_et_de_prospection_amortissement + \
            brevets_licence_logiciel_et_droits_similaires_amortissement + \
            fonds_commercial_et_droit_au_bail_amortissement + \
            autre_immobilisation_corporelles_amortissement

        # immobilisation_corporel
        immobilisation_corporel = terrains + batiment + amenagements_agencements_installations + \
            materiel_mobilier_et_actifs_biologiques + materiel_de_transport + \
            avance_et_acomptes_verse_sur_immobilisation

        # amortissement
        immobilisation_corporel_amortissement = terrains_amortissement + batiment_amortissement + amenagements_agencements_installations_amortissement + \
            materiel_mobilier_et_actifs_biologiques_amortissement + materiel_de_transport_amortissement + \
            avance_et_acomptes_verse_sur_immobilisation_amortissement

        # immobilisation financière
        immobilisation_financiere = titre_de_participation + \
            autres_immobilisation_financiere

        immobilisation_financiere_amortissement = titre_de_participation_amortissement + \
            autres_immobilisation_financiere_amortissement

        # total actif immobilise
        total_actif_immobilise = immobilisation_incorporel + \
            immobilisation_corporel + immobilisation_financiere
        # amortissement
        total_actif_immobilise_amortissement = immobilisation_incorporel_amortissement + \
            immobilisation_corporel_amortissement + immobilisation_financiere_amortissement

        total_actif_cirulant = actif_circulant_hao + stock_et_en_cours + \
            fournisseurs_avances_versees+clients+autres_creances

        # amortissement
        total_actif_cirulant_amortissement = actif_circulant_hao_amortissement + stock_et_en_cours_amortissement + \
            fournisseurs_avances_versees_amortissement + \
            clients_amortissement + autres_creances_amortissement
        # total_general
        total_general = total_actif_immobilise + \
            total_actif_cirulant + total_tresorerie_actif

        # Amortissements
        total_general_amortissement = total_actif_immobilise_amortissement + \
            total_actif_cirulant_amortissement + total_tresorerie_actif_amortissement
        # Total CAPRO RA
        total_capitaux_propres_et_ressources_assimilees = capital + apporteurs_captial_non_appele + prime_liees_au_capital_social + ecarts_de_reevaluation + \
            reserve_indisponible + reserve_libre + report_a_nouveau + \
            resultat_net_de_exercice + subvention_investissement + provision_reglementees

        # TOTAL DEFIRA
        total_dettes_financieres_et_ressources_assimilees = emprunts_et_dettes_financieres_divers + \
            dettes_de_location_acquisition + provisions_pour_risques_et_charges

        # Total resource stable
        total_ressources_stables = total_capitaux_propres_et_ressources_assimilees + \
            total_dettes_financieres_et_ressources_assimilees

        # Total passif circulant
        total_passif_circulant = dettes_circulantes_hao + clients_avance_recues + fournisseur_dexploitation + \
            dettes_fiscales_et_sociales + autres_dettes + provision_pour_risque_a_court_terme

        # total tresorerie passif
        total_tresorerie_passif = banque_credit_d_escompte + \
            banque_etablissement_financier_et_credit_de_tresorerie

        # total general passif
        total_general_passif = total_ressources_stables + total_passif_circulant + \
            total_tresorerie_passif + ecart_de_converion_passif
        # date immobilisation_incorporel
        data['immobilisation_incorporel'] = immobilisation_incorporel
        data['immobilisation_incorporel_amortissement'] = immobilisation_incorporel_amortissement
        data['immobilisation_incorporel_net'] = immobilisation_incorporel + \
            immobilisation_incorporel_amortissement

        data['frais_de_developpement_et_de_prospection'] = frais_de_developpement_et_de_prospection
        data['frais_de_developpement_et_de_prospection_amortissement'] = frais_de_developpement_et_de_prospection_amortissement
        data['frais_de_developpement_et_de_prospection_net'] = frais_de_developpement_et_de_prospection + \
            frais_de_developpement_et_de_prospection_amortissement

        data['brevets_licence_logiciel_et_droits_similaires'] = brevets_licence_logiciel_et_droits_similaires
        data['brevets_licence_logiciel_et_droits_similaires_amortissement'] = brevets_licence_logiciel_et_droits_similaires_amortissement
        data['brevets_licence_logiciel_et_droits_similaires_net'] = brevets_licence_logiciel_et_droits_similaires + \
            brevets_licence_logiciel_et_droits_similaires_amortissement

        data['fonds_commercial_et_droit_au_bail'] = fonds_commercial_et_droit_au_bail
        data['fonds_commercial_et_droit_au_bail_amortissement'] = fonds_commercial_et_droit_au_bail_amortissement
        data['fonds_commercial_et_droit_au_bail_net'] = fonds_commercial_et_droit_au_bail + \
            fonds_commercial_et_droit_au_bail_amortissement

        data['autre_immobilisation_corporelles'] = autre_immobilisation_corporelles
        data['autre_immobilisation_corporelles_amortissement'] = autre_immobilisation_corporelles_amortissement
        data['autre_immobilisation_corporelles_net'] = autre_immobilisation_corporelles + \
            autre_immobilisation_corporelles_amortissement

        # data immobilisation_corporel
        data['immobilisation_corporel'] = immobilisation_corporel
        data['immobilisation_corporel_amortissement'] = immobilisation_corporel_amortissement
        data['immobilisation_corporel_net'] = immobilisation_corporel + \
            immobilisation_corporel_amortissement

        data['terrains'] = terrains
        data['terrains_amortissement'] = terrains_amortissement
        data['terrains_net'] = terrains + terrains_amortissement

        data['batiment'] = batiment
        data['batiment_amortissement'] = batiment_amortissement
        data['batiment_net'] = batiment + batiment_amortissement

        data['amenagements_agencements_installations'] = amenagements_agencements_installations
        data['amenagements_agencements_installations_amortissement'] = amenagements_agencements_installations_amortissement
        data['amenagements_agencements_installations_net'] = amenagements_agencements_installations + \
            amenagements_agencements_installations_amortissement

        data['materiel_mobilier_et_actifs_biologiques'] = materiel_mobilier_et_actifs_biologiques
        data['materiel_mobilier_et_actifs_biologiques_amortissement'] = materiel_mobilier_et_actifs_biologiques_amortissement
        data['materiel_mobilier_et_actifs_biologiques_net'] = materiel_mobilier_et_actifs_biologiques + \
            materiel_mobilier_et_actifs_biologiques_amortissement

        data['materiel_de_transport'] = materiel_de_transport
        data['materiel_de_transport_amortissement'] = materiel_de_transport_amortissement
        data['materiel_de_transport_net'] = materiel_de_transport + \
            materiel_de_transport_amortissement

        data['avance_et_acomptes_verse_sur_immobilisation'] = avance_et_acomptes_verse_sur_immobilisation
        data['avance_et_acomptes_verse_sur_immobilisation_amortissement'] = avance_et_acomptes_verse_sur_immobilisation_amortissement
        data['avance_et_acomptes_verse_sur_immobilisation_net'] = avance_et_acomptes_verse_sur_immobilisation + \
            avance_et_acomptes_verse_sur_immobilisation_amortissement

        # immobilisation financière
        data['immobilisation_financiere'] = immobilisation_financiere
        data['immobilisation_financiere_amortissement'] = immobilisation_financiere_amortissement
        data['immobilisation_financiere_net'] = immobilisation_financiere + \
            immobilisation_financiere_amortissement

        data['titre_de_participation'] = titre_de_participation
        data['titre_de_participation_amortissement'] = titre_de_participation_amortissement
        data['titre_de_participation_net'] = titre_de_participation + \
            titre_de_participation_amortissement

        data['autres_immobilisation_financiere'] = autres_immobilisation_financiere
        data['autres_immobilisation_financiere_amortissement'] = autres_immobilisation_financiere_amortissement
        data['autres_immobilisation_financiere_net'] = autres_immobilisation_financiere + \
            autres_immobilisation_financiere_amortissement

        data['total_actif_immobilise'] = total_actif_immobilise
        data['total_actif_immobilise_amortissement'] = total_actif_immobilise_amortissement
        data['total_actif_immobilise_net'] = total_actif_immobilise + \
            total_actif_immobilise_amortissement

        # actif ciculant
        data['actif_circulant_hao'] = actif_circulant_hao
        data['actif_circulant_hao_amortissement'] = actif_circulant_hao_amortissement
        data['actif_circulant_hao_net'] = actif_circulant_hao + \
            actif_circulant_hao_amortissement

        data['stock_et_en_cours'] = stock_et_en_cours
        data['stock_et_en_cours_amortissement'] = stock_et_en_cours_amortissement
        data['stock_et_en_cours_net'] = stock_et_en_cours + \
            stock_et_en_cours_amortissement

        data['creances_et_emplois_assimiles'] = creances_et_emplois_assimiles
        data['creances_et_emplois_assimiles_amortissement'] = creances_et_emplois_assimiles_amortissement
        data['creances_et_emplois_assimiles_net'] = creances_et_emplois_assimiles + \
            creances_et_emplois_assimiles_amortissement

        data['fournisseurs_avances_versees'] = fournisseurs_avances_versees
        data['fournisseurs_avances_versees_amortissement'] = fournisseurs_avances_versees_amortissement
        data['fournisseurs_avances_versees_net'] = fournisseurs_avances_versees + \
            fournisseurs_avances_versees_amortissement

        data['clients'] = clients
        data['clients_amortissement'] = clients_amortissement
        data['clients_net'] = clients + clients_amortissement

        data['autres_creances'] = autres_creances
        data['autres_creances_amortissement'] = autres_creances_amortissement
        data['autres_creances_net'] = autres_creances + \
            autres_creances_amortissement

        data['total_actif_cirulant'] = total_actif_cirulant
        data['total_actif_cirulant_amortissement'] = total_actif_cirulant_amortissement
        data['total_actif_cirulant_net'] = total_actif_cirulant + \
            total_actif_cirulant_amortissement

        # tresorerie actif
        data['titres_de_placement'] = titres_de_placement
        data['titres_de_placement_amortissement'] = titres_de_placement_amortissement
        data['titres_de_placement_net'] = titres_de_placement + \
            titres_de_placement_amortissement

        data['valeur_a_encaisser'] = valeur_a_encaisser
        data['valeur_a_encaisser_amortissement'] = valeur_a_encaisser_amortissement
        data['valeur_a_encaisser_net'] = valeur_a_encaisser + \
            valeur_a_encaisser_amortissement

        data['banque_cheque_postaux_caisse_et_assimile'] = banque_cheque_postaux_caisse_et_assimile
        data['banque_cheque_postaux_caisse_et_assimile_amortissement'] = banque_cheque_postaux_caisse_et_assimile_amortissement
        data['banque_cheque_postaux_caisse_et_assimile_net'] = banque_cheque_postaux_caisse_et_assimile + \
            banque_cheque_postaux_caisse_et_assimile_amortissement

        data['total_tresorerie_actif'] = total_tresorerie_actif
        data['total_tresorerie_actif_amortissement'] = total_tresorerie_actif_amortissement
        data['total_tresorerie_actif_net'] = total_tresorerie_actif + \
            total_tresorerie_actif_amortissement

        # ecart_de_conversion_actif
        data['ecart_de_conversion_actif'] = ecart_de_conversion_actif

        data['total_general'] = total_general
        data['total_general_amortissement'] = total_general_amortissement
        data['total_general_net'] = total_general + total_general_amortissement

        # PASSIF DU BILAN
        # CAPRO RA
        data['capital'] = capital
        data['apporteurs_captial_non_appele'] = apporteurs_captial_non_appele
        data['prime_liees_au_capital_social'] = prime_liees_au_capital_social
        data['ecarts_de_reevaluation'] = ecarts_de_reevaluation
        data['reserve_indisponible'] = reserve_indisponible
        data['reserve_libre'] = reserve_libre
        data['report_a_nouveau'] = report_a_nouveau
        data['resultat_net_de_exercice'] = resultat_net_de_exercice
        data['subvention_investissement'] = subvention_investissement
        data['provision_reglementees'] = provision_reglementees
        data['total_capitaux_propres_et_ressources_assimilees'] = total_capitaux_propres_et_ressources_assimilees

        # DEFI RA
        data['emprunts_et_dettes_financieres_divers'] = emprunts_et_dettes_financieres_divers
        data['dettes_de_location_acquisition'] = dettes_de_location_acquisition
        data['provisions_pour_risques_et_charges'] = provisions_pour_risques_et_charges
        data['total_dettes_financieres_et_ressources_assimilees'] = total_dettes_financieres_et_ressources_assimilees
        data['total_ressources_stables'] = total_ressources_stables

        # PASSIF CIRCULANT
        data['dettes_circulantes_hao'] = dettes_circulantes_hao
        data['clients_avance_recues'] = clients_avance_recues
        data['fournisseur_dexploitation'] = fournisseur_dexploitation
        data['dettes_fiscales_et_sociales'] = dettes_fiscales_et_sociales
        data['autres_dettes'] = autres_dettes
        data['provision_pour_risque_a_court_terme'] = provision_pour_risque_a_court_terme
        data['total_passif_circulant'] = total_passif_circulant
        data['banque_credit_d_escompte'] = banque_credit_d_escompte
        data['banque_etablissement_financier_et_credit_de_tresorerie'] = banque_etablissement_financier_et_credit_de_tresorerie
        data['total_tresorerie_passif'] = total_tresorerie_passif
        data['ecart_de_converion_passif'] = ecart_de_converion_passif

        data['total_general_passif'] = total_general_passif

        data['excercice'] = self.date_fin.year
        return data

    def get_item_data(self):
        file_name = _('Bilan ohada.xlsx')
        fp = BytesIO()
        data = self.get_bilan_data()
        workbook = xlsxwriter.Workbook(fp)
        worksheet = workbook.add_worksheet("BILAN ACTIF")
        worksheet.write(2, 0, "REF")
        worksheet.write(2, 1, "ACTIF")
        worksheet.write(2, 2, "BRUT")
        worksheet.write(2, 3, "AMMORTISSEMENT")
        worksheet.write(2, 4, "NET")
        cell_number_format = workbook.add_format(
            {'align': 'right', 'bold': False, 'size': 12, 'num_format': '#,###0'})
        code_list = ["AD", "AE", "AF", "AG",
                     "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AP", "AQ", "AR", "AS", "AZ", "BA", "BB", "BG", "BH", "BI", "BJ", "BK", "BQ", "BR", "BS", "BT", "BU", "BZ"]
        linge = 3
        for code in code_list:
            worksheet.write(code_list.index(code) + linge, 0, code)

        postes_list = ['IMMOBILISATIONS INCORPORELLES',
                       'Frais de développement et de prospection', 'Brevets, licences, logiciels et droits similaires', 'Fonds commercial et droit au bail', 'Autres immobilisations incorporelles', 'IMMOBILISATIONS CORPORELLES', 'Terrains (Dont Plancement en net)', 'Bâtiment (Dont Plancement en net)', 'Aménagements, agencements et installations', 'Matériel, mobilier et actifs biologiques', 'Matériel de transport', 'AVANCES ET ACOMPTES VERSES SUR IMMOBILISATIONS', 'IMMOBILISATIONS FINANCIERES', 'Titres de participation', 'Autres immobilisations financières', 'TOTAL ACTIF IMMOBILISE', 'ACTIF CIRCULANT HAO', 'STOCKS ET ENCOURS', 'CREANCES ET EMPLOIS ASSIMILES', 'Fournisseurs avances versées', 'Clients', 'Autres créances', 'TOTAL ACTIF CIRCULANT', 'Titres de placement', 'Valeurs à encaisser', 'Banques, chèques postaux, caisse et assimilés', 'TOTAL TRESORERIE ACTIF', 'Ecart de conversion-Actif', 'TOTAL GENERAL']

        for poste in postes_list:
            worksheet.write(postes_list.index(poste) + linge, 1, poste)
        worksheet.merge_range("A1:D1", 'BILAN SYSCOADA REVISE (ACTIF)')
        worksheet.set_column('B:B', 55)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('D:D', 20)
        worksheet.set_column('E:E', 20)

        worksheet.write(
            3, 2, data['immobilisation_incorporel'], cell_number_format)
        worksheet.write(
            3, 3, data['immobilisation_incorporel_amortissement'], cell_number_format)
        worksheet.write(
            3, 4, data['immobilisation_incorporel_net'], cell_number_format)

        worksheet.write(
            4, 2, data['frais_de_developpement_et_de_prospection'], cell_number_format)
        worksheet.write(
            4, 3, data['frais_de_developpement_et_de_prospection_amortissement'], cell_number_format)
        worksheet.write(
            4, 4, data['frais_de_developpement_et_de_prospection_net'], cell_number_format)

        worksheet.write(
            5, 2, data['brevets_licence_logiciel_et_droits_similaires'], cell_number_format)
        worksheet.write(
            5, 3, data['brevets_licence_logiciel_et_droits_similaires_amortissement'], cell_number_format)
        worksheet.write(
            5, 4, data['brevets_licence_logiciel_et_droits_similaires_net'], cell_number_format)

        worksheet.write(
            6, 2, data['fonds_commercial_et_droit_au_bail'], cell_number_format)
        worksheet.write(
            6, 3, data['fonds_commercial_et_droit_au_bail_amortissement'], cell_number_format)
        worksheet.write(
            6, 4, data['fonds_commercial_et_droit_au_bail_net'], cell_number_format)

        worksheet.write(
            7, 2, data['autre_immobilisation_corporelles'], cell_number_format)
        worksheet.write(
            7, 3, data['autre_immobilisation_corporelles_amortissement'], cell_number_format)
        worksheet.write(
            7, 4, data['autre_immobilisation_corporelles_net'], cell_number_format)

        worksheet.write(
            8, 2, data['immobilisation_corporel'], cell_number_format)
        worksheet.write(
            8, 3, data['immobilisation_corporel_amortissement'], cell_number_format)
        worksheet.write(
            8, 4, data['immobilisation_corporel_net'], cell_number_format)

        worksheet.write(9, 2, data['terrains'], cell_number_format)
        worksheet.write(
            9, 3, data['terrains_amortissement'], cell_number_format)
        worksheet.write(9, 4, data['terrains_net'], cell_number_format)

        worksheet.write(10, 2, data['batiment'], cell_number_format)
        worksheet.write(
            10, 3, data['batiment_amortissement'], cell_number_format)
        worksheet.write(10, 4, data['batiment_net'], cell_number_format)

        worksheet.write(
            11, 2, data['amenagements_agencements_installations'], cell_number_format)
        worksheet.write(
            11, 3, data['amenagements_agencements_installations_amortissement'], cell_number_format)
        worksheet.write(
            11, 4, data['amenagements_agencements_installations_net'], cell_number_format)

        worksheet.write(
            12, 2, data['materiel_mobilier_et_actifs_biologiques'], cell_number_format)
        worksheet.write(
            12, 3, data['materiel_mobilier_et_actifs_biologiques_amortissement'], cell_number_format)
        worksheet.write(
            12, 4, data['materiel_mobilier_et_actifs_biologiques_net'], cell_number_format)

        worksheet.write(
            13, 2, data['materiel_de_transport'], cell_number_format)
        worksheet.write(
            13, 3, data['materiel_de_transport_amortissement'], cell_number_format)
        worksheet.write(
            13, 4, data['materiel_de_transport_net'], cell_number_format)

        worksheet.write(
            14, 2, data['avance_et_acomptes_verse_sur_immobilisation'], cell_number_format)
        worksheet.write(
            14, 3, data['avance_et_acomptes_verse_sur_immobilisation_amortissement'], cell_number_format)
        worksheet.write(
            14, 4, data['avance_et_acomptes_verse_sur_immobilisation_net'], cell_number_format)

        worksheet.write(
            15, 2, data['immobilisation_financiere'], cell_number_format)
        worksheet.write(
            15, 3, data['immobilisation_financiere_amortissement'], cell_number_format)
        worksheet.write(
            15, 4, data['immobilisation_financiere_net'], cell_number_format)

        worksheet.write(
            16, 2, data['titre_de_participation'], cell_number_format)
        worksheet.write(
            16, 3, data['titre_de_participation_amortissement'], cell_number_format)
        worksheet.write(
            16, 4, data['titre_de_participation_net'], cell_number_format)

        worksheet.write(
            17, 2, data['autres_immobilisation_financiere'], cell_number_format)
        worksheet.write(
            17, 3, data['autres_immobilisation_financiere_amortissement'], cell_number_format)
        worksheet.write(
            17, 4, data['autres_immobilisation_financiere_net'], cell_number_format)

        worksheet.write(
            18, 2, data['total_actif_immobilise'], cell_number_format)
        worksheet.write(
            18, 3, data['total_actif_immobilise_amortissement'], cell_number_format)
        worksheet.write(
            18, 4, data['total_actif_immobilise_net'], cell_number_format)

        worksheet.write(
            19, 2, data['actif_circulant_hao'], cell_number_format)
        worksheet.write(
            19, 3, data['actif_circulant_hao_amortissement'], cell_number_format)
        worksheet.write(
            19, 4, data['actif_circulant_hao_net'], cell_number_format)

        worksheet.write(
            20, 2, data['stock_et_en_cours'], cell_number_format)
        worksheet.write(
            20, 3, data['stock_et_en_cours_amortissement'], cell_number_format)
        worksheet.write(
            20, 4, data['stock_et_en_cours_net'], cell_number_format)

        worksheet.write(
            21, 2, data['creances_et_emplois_assimiles'], cell_number_format)
        worksheet.write(
            21, 3, data['creances_et_emplois_assimiles_amortissement'], cell_number_format)
        worksheet.write(
            21, 4, data['creances_et_emplois_assimiles_net'], cell_number_format)

        worksheet.write(
            22, 2, data['fournisseurs_avances_versees'], cell_number_format)
        worksheet.write(
            22, 3, data['fournisseurs_avances_versees_amortissement'], cell_number_format)
        worksheet.write(
            22, 4, data['fournisseurs_avances_versees_net'], cell_number_format)

        worksheet.write(
            23, 2, data['clients'], cell_number_format)
        worksheet.write(
            23, 3, data['clients_amortissement'], cell_number_format)
        worksheet.write(
            23, 4, data['clients_net'], cell_number_format)

        worksheet.write(
            24, 2, data['autres_creances'], cell_number_format)
        worksheet.write(
            24, 3, data['autres_creances_amortissement'], cell_number_format)
        worksheet.write(
            24, 4, data['autres_creances_net'], cell_number_format)

        worksheet.write(
            25, 2, data['total_actif_cirulant'], cell_number_format)
        worksheet.write(
            25, 3, data['total_actif_cirulant_amortissement'], cell_number_format)
        worksheet.write(
            25, 4, data['total_actif_cirulant_net'], cell_number_format)

        worksheet.write(26, 2, data['titres_de_placement'], cell_number_format)
        worksheet.write(
            26, 3, data['titres_de_placement_amortissement'], cell_number_format)
        worksheet.write(
            26, 4, data['titres_de_placement_net'], cell_number_format)

        worksheet.write(27, 2, data['valeur_a_encaisser'], cell_number_format)
        worksheet.write(
            27, 3, data['valeur_a_encaisser_amortissement'], cell_number_format)
        worksheet.write(
            27, 4, data['valeur_a_encaisser_net'], cell_number_format)

        worksheet.write(
            28, 2, data['banque_cheque_postaux_caisse_et_assimile'], cell_number_format)
        worksheet.write(
            28, 3, data['banque_cheque_postaux_caisse_et_assimile_amortissement'], cell_number_format)
        worksheet.write(
            28, 4, data['banque_cheque_postaux_caisse_et_assimile_net'], cell_number_format)

        worksheet.write(
            29, 2, data['total_tresorerie_actif'], cell_number_format)
        worksheet.write(
            29, 3, data['total_tresorerie_actif_amortissement'], cell_number_format)
        worksheet.write(
            29, 4, data['total_tresorerie_actif_net'], cell_number_format)

        worksheet.write(
            31, 3, data['ecart_de_conversion_actif'], cell_number_format)

        worksheet.write(31, 2, data['total_general'], cell_number_format)
        worksheet.write(
            31, 3, data['total_general_amortissement'], cell_number_format)
        worksheet.write(31, 4, data['total_general_net'], cell_number_format)

        worksheet = workbook.add_worksheet("BILAN PASSIF")
        worksheet.merge_range("A1:D1", 'BILAN SYSCOADA REVISE (PASSIF)')
        worksheet.write(2, 0, "REF")
        worksheet.write(2, 1, "PASSIF")
        worksheet.write(2, 2, "NOTE")
        worksheet.write(2, 3, "NET")
        worksheet.set_column('B:B', 55)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('D:D', 20)
        code_list = ["CA", "CB", "CD", "CE",
                     "CF", "CG", "CH", "CJ", "CL", "CM", "CP", "DA", "DB", "DC", "DD", "DF", "DH", "DI", "DJ", "DK", "DM", "DN", "DP", "DQ", "DR", "DT", "DV", "DZ"]
        linge = 3
        for code in code_list:
            worksheet.write(code_list.index(code) + linge, 0, code)

        post_list = ['Capital', 'Apporteurs capital non appelé', 'Primes liées au capital social', 'Ecarts de réévaluation', 'Réserves indisponibles', 'Réserves libres', 'Report à nouveau', 'Résultat net de l\'exercice', 'Subventions d\'investissement', 'Provisions réglementées', 'TOTAL CAPITAUX PROPRES ET RESSOURCES ASSIMILEES', 'Emprunts et dettes financières diverses', 'Dettes de location acquisition', 'Provisions pour risques et charges',
                     'TOTAL DETTES FINANCIERES ET RESSOURCES ASSIMILEES', 'TOTAL RESSOURCES STABLES', 'Dettes circulantes HAO', 'Clients, avances reçues', 'Fournisseurs d\'exploitation', 'Dettes fiscales et sociales', 'Autres dettes', 'Provisions pour risques à court terme', 'TOTAL PASSIF CIRCULANT', 'Banques, crédits d\'escompte', "Banques, établissements financiers et crédits de trésorerie", 'TOTAL TRESORERIE PASSIF', 'Ecart de conversion-Passif', 'TOTAL GENERAL']
        linge = 3
        for code in post_list:
            worksheet.write(post_list.index(code) + linge, 1, code)

        worksheet.write(3, 3, -data['capital'], cell_number_format)
        worksheet.write(
            4, 3, -data['apporteurs_captial_non_appele'], cell_number_format)
        worksheet.write(
            5, 3, -data['prime_liees_au_capital_social'], cell_number_format)
        worksheet.write(
            6, 3, -data['ecarts_de_reevaluation'], cell_number_format)
        worksheet.write(
            7, 3, -data['reserve_indisponible'], cell_number_format)
        worksheet.write(8, 3, -data['reserve_libre'], cell_number_format)
        worksheet.write(9, 3, -data['report_a_nouveau'], cell_number_format)
        worksheet.write(
            10, 3, -data['resultat_net_de_exercice'], cell_number_format)
        worksheet.write(
            11, 3, -data['subvention_investissement'], cell_number_format)
        worksheet.write(
            12, 3, -data['provision_reglementees'], cell_number_format)
        worksheet.write(
            13, 3, -data['total_capitaux_propres_et_ressources_assimilees'], cell_number_format)
        worksheet.write(
            14, 3, -data['emprunts_et_dettes_financieres_divers'], cell_number_format)
        worksheet.write(
            15, 3, -data['dettes_de_location_acquisition'], cell_number_format)
        worksheet.write(
            16, 3, -data['provisions_pour_risques_et_charges'], cell_number_format)
        worksheet.write(
            17, 3, -data['total_dettes_financieres_et_ressources_assimilees'], cell_number_format)
        worksheet.write(
            18, 3, -data['total_ressources_stables'], cell_number_format)
        worksheet.write(
            19, 3, -data['dettes_circulantes_hao'], cell_number_format)
        worksheet.write(
            20, 3, -data['clients_avance_recues'], cell_number_format)
        worksheet.write(
            21, 3, -data['fournisseur_dexploitation'], cell_number_format)
        worksheet.write(
            22, 3, -data['dettes_fiscales_et_sociales'], cell_number_format)
        worksheet.write(23, 3, -data['autres_dettes'], cell_number_format)
        worksheet.write(
            24, 3, -data['provision_pour_risque_a_court_terme'], cell_number_format)
        worksheet.write(
            25, 3, -data['total_passif_circulant'], cell_number_format)
        worksheet.write(
            26, 3, -data['banque_credit_d_escompte'], cell_number_format)
        worksheet.write(
            27, 3, -data['banque_etablissement_financier_et_credit_de_tresorerie'], cell_number_format)
        worksheet.write(
            28, 3, -data['total_tresorerie_passif'], cell_number_format)
        worksheet.write(
            29, 3, -data['ecart_de_converion_passif'], cell_number_format)
        worksheet.write(
            30, 3, -data['total_general_passif'], cell_number_format)

        workbook.close()
        file_download = base64.b64encode(fp.getvalue())
        fp.close()
        self = self.with_context(
            default_name=file_name, default_file_download=file_download)

        return {
            'name': 'Téléchargement du bilan',
            'view_mode': 'form',
            'res_model': 'bilan.excel',
            'type': 'ir.actions.act_window',
            'target': 'new',
            'context': self._context,
        }


class bilan_excel(models.TransientModel):
    _name = 'bilan.excel'

    name = fields.Char('Nom du fichier', size=256, readonly=True)
    file_download = fields.Binary('Télécharger le bilan', readonly=True)
