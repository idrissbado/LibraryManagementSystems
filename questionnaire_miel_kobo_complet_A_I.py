# XLSForm generator for KoboCollect - Questionnaire Commerçants Miel Côte d'Ivoire
# This script creates an XLSForm Excel file with all modules and instructions as specified
import xlsxwriter

# Create a workbook and add worksheets
workbook = xlsxwriter.Workbook('questionnaire_miel_kobo_commercants.xlsx')
survey = workbook.add_worksheet('survey')
choices = workbook.add_worksheet('choices')

# Survey sheet headers
survey_headers = [
    'type', 'name', 'label', 'appearance', 'relevant', 'repeat_count'
]
survey.write_row(0, 0, survey_headers)

# Survey sheet rows (full structure)
survey_rows = [
    ['note', 'intro', '🧭 QUESTIONNAIRE COMMERÇANTS\nÉTUDE DE LA FILIÈRE MIEL EN CÔTE D’IVOIRE\nObjectif général : Collecter des données fiables et non redondantes auprès des commerçants de miel, afin d’analyser les circuits de commercialisation, la formation des prix, la qualité, les contraintes, les marges et de simuler des scénarios de chocs et de développement du marché.\nInstructions à l’enquêteur : Cocher les cases correspondantes. Renseigner précisément les champs numériques et textuels. Les sections « par période » sont à répéter si nécessaire.', '', '', ''],
    ['begin_group', 'module_a', 'MODULE A – IDENTIFICATION ET LOCALISATION', '', '', ''],
    ['text', 'enqueteur', 'Nom et prénom de l’enquêteur', '', '', ''],
    ['date', 'date_entretien', 'Date de l’entretien', '', '', ''],
    ['select_one region', 'region', 'Région', '', '', ''],
    ['text', 'departement', 'Département', '', '', ''],
    ['text', 'localite', 'Commune / Localité', '', '', ''],
    ['select_one vente_lieu', 'type_vente', 'Type de lieu de vente', '', '', ''],
    ['end_group', '', '', '', '', ''],
    ['begin_group', 'module_b', 'MODULE B – PROFIL DU COMMERÇANT', '', '', ''],
    ['select_one categorie', 'categorie', 'Catégorie', '', '', ''],
    ['select_one yes_no', 'membre_org', 'Êtes-vous membre d’une organisation / association de commerçants ?', '', '', ''],
    ['text', 'nom_structure', 'Si oui, nom de la structure', '', "${membre_org}='yes'", ''],
    ['integer', 'age', 'Âge (ans)', '', '', ''],
    ['select_one sexe', 'sexe', 'Sexe', '', '', ''],
    ['text', 'nationalite', 'Nationalité', '', '', ''],
    ['text', 'telephone', 'Téléphone', '', '', ''],
    ['select_one education', 'education', 'Niveau d’éducation', '', '', ''],
    ['select_one marital', 'statut_matrimonial', 'Statut matrimonial', '', '', ''],
    ['end_group', '', '', '', '', ''],
    ['begin_group', 'module_c', 'MODULE C – ACTIVITÉ COMMERCIALE DU MIEL', '', '', ''],
    ['integer', 'annee_experience', 'Années d’expérience dans le commerce du miel', '', '', ''],
    ['select_multiple appro', 'approvisionnement', 'Principales sources d’approvisionnement', '', '', ''],
    ['select_one origine', 'origine_miel', 'Origine géographique du miel', '', '', ''],
    ['select_one yes_no', 'appro_regulier', 'Approvisionnement régulier toute l’année', '', '', ''],
    ['text', 'periode_forte', 'Périodes de forte disponibilité', '', '', ''],
    ['end_group', '', '', '', '', ''],
    ['begin_group', 'module_d', 'MODULE D – VOLUMES ET FLUX COMMERCIAUX', '', '', ''],
    ['integer', 'volume_achete', 'Volume moyen acheté par période (kg)', '', '', ''],
    ['integer', 'volume_vendu', 'Volume moyen vendu par période (kg)', '', '', ''],
    ['select_multiple miel', 'type_miel', 'Types de miel commercialisés', '', '', ''],
    ['select_one forme', 'forme_vente', 'Forme de vente', '', '', ''],
    ['end_group', '', '', '', '', ''],
    ['begin_group', 'module_e', 'MODULE E – PRIX, COÛTS ET MARGES', '', '', ''],
    ['integer', 'prix_achat', 'Prix moyen d’achat (FCFA/kg)', '', '', ''],
    ['integer', 'prix_vente', 'Prix moyen de vente (FCFA/kg)', '', '', ''],
    ['select_multiple couts', 'couts_activite', 'Coûts liés à l’activité', '', '', ''],
    ['integer', 'marge_brute', 'Marge brute estimée (FCFA/kg)', '', '', ''],
    ['end_group', '', '', '', '', ''],
    ['begin_group', 'module_f', 'MODULE F – QUALITÉ, CONSERVATION ET CONFORMITÉ', '', '', ''],
    ['select_multiple qual', 'criteres_qualite', 'Critères de qualité à l’achat', '', '', ''],
    ['select_multiple stockage', 'conditions_stockage', 'Conditions de stockage', '', '', ''],
    ['select_one contenant', 'type_contenant', 'Type de contenant', '', '', ''],
    ['select_multiple prob_qual', 'problemes_qualite', 'Problèmes de qualité rencontrés', '', '', ''],
    ['end_group', '', '', '', '', ''],
    ['begin_group', 'module_g', 'MODULE G – PERTES ET RISQUES COMMERCIAUX', '', '', ''],
    ['select_one yes_no', 'pertes', 'Subissez-vous des pertes de miel ?', '', '', ''],
    ['integer', 'taux_pertes', 'Taux estimé de pertes (%)', '', "${pertes}='yes'", ''],
    ['select_multiple causes', 'causes_pertes', 'Causes des pertes', '', "${pertes}='yes'", ''],
    ['end_group', '', '', '', '', ''],
    ['begin_group', 'module_h', 'MODULE H – MARCHÉS ET CLIENTÈLE', '', '', ''],
    ['select_multiple clients', 'types_clients', 'Types de clients', '', '', ''],
    ['select_multiple marche', 'portee_marche', 'Portée du marché', '', '', ''],
    ['select_multiple prix', 'fixation_prix', 'Fixation des prix', '', '', ''],
    ['end_group', '', '', '', '', ''],
    ['begin_group', 'module_i', 'MODULE I – APPUIS ET BESOINS', '', '', ''],
    ['select_multiple appui', 'appuis_recus', 'Appuis reçus', '', '', ''],
    ['select_multiple besoins', 'besoins_prioritaires', 'Besoins prioritaires', '', '', ''],
    ['end_group', '', '', '', '', ''],
    ['begin_group', 'module_j', 'MODULE J – STRESS TEST ET SCÉNARIOS DE MARCHÉ', '', '', ''],
    ['integer', 'volume_annuel', 'Volume annuel vendu (kg/an)', '', '', ''],
    ['integer', 'ca_annuel', 'Chiffre d’affaires annuel estimé (FCFA)', '', '', ''],
    ['select_one amelior_qual', 'scenar_qualite', 'Amélioration de la qualité du miel', '', '', ''],
    ['select_one orga_com', 'scenar_orga', 'Meilleure organisation des commerçants', '', '', ''],
    ['select_one financement', 'scenar_financement', 'Accès au financement', '', '', ''],
    ['select_one hausse_prix', 'scenar_hausse_prix', 'Hausse des prix d’achat', '', '', ''],
    ['select_one baisse_dem', 'scenar_baisse_dem', 'Baisse de la demande', '', '', ''],
    ['select_one prob_qual_gen', 'scenar_prob_qualite', 'Problèmes de qualité généralisés', '', '', ''],
    ['select_one facteur', 'facteur_cle', 'Quel facteur aurait le plus grand impact sur votre activité ?', '', '', ''],
    ['end_group', '', '', '', '', ''],
    ['begin_group', 'module_k', 'MODULE K – OBSERVATIONS', '', '', ''],
    ['text', 'commentaires', 'Commentaires du commerçant', '', '', ''],
    ['text', 'observations', 'Observations de l’enquêteur', '', '', ''],
    ['end_group', '', '', '', '', ''],
]

for i, row in enumerate(survey_rows, 1):
    survey.write_row(i, 0, row)

choices_headers = ['list_name', 'name', 'label']
choices.write_row(0, 0, choices_headers)

choices_rows = [
    # Regions (add all as needed)
    ['region', 'abidjan', 'Abidjan'],
    ['region', 'bouake', 'Bouaké'],
    ['region', 'yamoussoukro', 'Yamoussoukro'],
    # ... (add all regions of Côte d’Ivoire)
    ['vente_lieu', 'marche', 'Marché'],
    ['vente_lieu', 'boutique', 'Boutique'],
    ['vente_lieu', 'domicile', 'Domicile'],
    ['vente_lieu', 'autre', 'Autre'],
    ['categorie', 'detail', 'Détaillant'],
    ['categorie', 'grossiste', 'Grossiste'],
    ['categorie', 'semi', 'Semi-grossiste'],
    ['categorie', 'collecteur', 'Collecteur–revendeur'],
    ['yes_no', 'yes', 'Oui'],
    ['yes_no', 'no', 'Non'],
    ['sexe', 'femme', 'Femme'],
    ['sexe', 'homme', 'Homme'],
    ['education', 'non_scol', 'Non scolarisé'],
    ['education', 'primaire', 'Primaire'],
    ['education', 'secondaire', 'Secondaire'],
    ['education', 'superieur', 'Supérieur'],
    ['marital', 'marie', 'Marié(e)'],
    ['marital', 'celibataire', 'Célibataire'],
    ['marital', 'concubinage', 'Concubinage'],
    ['appro', 'prod', 'Producteurs individuels'],
    ['appro', 'coop', 'Coopératives apicoles'],
    ['appro', 'collecteur', 'Collecteurs'],
    ['appro', 'commercant', 'Autres commerçants'],
    ['origine', 'locale', 'Locale'],
    ['origine', 'regionale', 'Régionale'],
    ['origine', 'nationale', 'Nationale'],
    ['origine', 'importee', 'Importée'],
    ['miel', 'foret', 'Forêt'],
    ['miel', 'savane', 'Savane'],
    ['miel', 'multifloral', 'Multifloral'],
    ['miel', 'autre', 'Autre'],
    ['forme', 'vrac', 'Vrac'],
    ['forme', 'conditionne', 'Conditionné'],
    ['forme', 'les_deux', 'Les deux'],
    ['couts', 'transport', 'Transport'],
    ['couts', 'stockage', 'Stockage'],
    ['couts', 'conditionnement', 'Conditionnement'],
    ['couts', 'taxes', 'Taxes / redevances'],
    ['couts', 'pertes', 'Pertes'],
    ['couts', 'autres', 'Autres'],
    ['qual', 'aspect', 'Aspect'],
    ['qual', 'gout', 'Goût'],
    ['qual', 'origine', 'Origine connue'],
    ['qual', 'confiance', 'Confiance au fournisseur'],
    ['qual', 'prix_bas', 'Prix bas'],
    ['stockage', 'sec', 'Sec'],
    ['stockage', 'abri_lumiere', 'À l’abri de la lumière'],
    ['stockage', 'hermetique', 'Récipients hermétiques'],
    ['contenant', 'plastique', 'Plastique alimentaire'],
    ['contenant', 'verre', 'Verre'],
    ['contenant', 'autre', 'Autre'],
    ['prob_qual', 'fermentation', 'Fermentation'],
    ['prob_qual', 'impuretes', 'Impuretés'],
    ['prob_qual', 'cristallisation', 'Cristallisation excessive'],
    ['prob_qual', 'adulteration', 'Adultération suspectée'],
    ['prob_qual', 'aucun', 'Aucun'],
    ['causes', 'conservation', 'Mauvaise conservation'],
    ['causes', 'chaleur', 'Chaleur excessive'],
    ['causes', 'contenant', 'Contenants inadaptés'],
    ['causes', 'mevente', 'Mévente'],
    ['causes', 'autre', 'Autre'],
    ['clients', 'menages', 'Ménages'],
    ['clients', 'revendeurs', 'Revendeurs'],
    ['clients', 'restaurateurs', 'Restaurateurs'],
    ['clients', 'pharmacies', 'Pharmacies / herboristeries'],
    ['marche', 'local', 'Local'],
    ['marche', 'regional', 'Régional'],
    ['marche', 'national', 'National'],
    ['prix', 'cout_marge', 'Coût + marge'],
    ['prix', 'marche', 'Prix du marché'],
    ['prix', 'concurrence', 'Concurrence'],
    ['prix', 'nego', 'Négociation avec clients'],
    ['appui', 'aucun', 'Aucun'],
    ['appui', 'formation', 'Formation'],
    ['appui', 'financement', 'Financement'],
    ['appui', 'organisation', 'Appui organisationnel'],
    ['besoins', 'financement', 'Accès au financement'],
    ['besoins', 'stockage', 'Stockage amélioré'],
    ['besoins', 'certification', 'Certification / contrôle qualité'],
    ['besoins', 'organisation', 'Organisation collective'],
    ['besoins', 'marche', 'Accès à de nouveaux marchés'],
    ['amelior_qual', 'plus10', '+10 %'],
    ['amelior_qual', 'plus25', '+25 %'],
    ['amelior_qual', 'plus50', '+50 %'],
    ['amelior_qual', 'plus50plus', '> +50 %'],
    ['orga_com', 'faible', 'Faible impact'],
    ['orga_com', 'moyen', 'Impact moyen'],
    ['orga_com', 'eleve', 'Impact élevé'],
    ['financement', 'volumes', 'Augmentation des volumes'],
    ['financement', 'couts', 'Réduction des coûts'],
    ['financement', 'les_deux', 'Les deux'],
    ['hausse_prix', 'moins10', '-10 %'],
    ['hausse_prix', 'moins25', '-25 %'],
    ['hausse_prix', 'moins50', '-50 %'],
    ['hausse_prix', 'moins50plus', '> -50 %'],
    ['baisse_dem', 'moins10', '-10 %'],
    ['baisse_dem', 'moins25', '-25 %'],
    ['baisse_dem', 'moins50', '-50 %'],
    ['baisse_dem', 'moins50plus', '> -50 %'],
    ['prob_qual_gen', 'faible', 'Impact faible'],
    ['prob_qual_gen', 'moyen', 'Impact moyen'],
    ['prob_qual_gen', 'eleve', 'Impact élevé'],
    ['facteur', 'qualite', 'Qualité du miel'],
    ['facteur', 'prix_achat', 'Prix d’achat'],
    ['facteur', 'organisation', 'Organisation des acteurs'],
    ['facteur', 'financement', 'Accès au financement'],
    ['facteur', 'demande', 'Demande des consommateurs'],
]

for i, row in enumerate(choices_rows, 1):
    choices.write_row(i, 0, row)

workbook.close()
print("XLSForm created: questionnaire_miel_kobo_commercants.xlsx")
import pandas as pd

# Onglet settings
def get_settings():
    return pd.DataFrame({
        'form_title': ['QUESTIONNAIRE PRODUCTEURS - ÉTUDE DE LA FILIÈRE MIEL EN CÔTE D’IVOIRE'],
        'form_id': ['miel_ci_complet_A_I'],
        'version': ['v1']
    })

# Onglet choices
def get_choices():
    return pd.DataFrame([
        # oui_non
        ['oui_non','oui','Oui'], ['oui_non','non','Non'],
        # sexe
        ['sexe','homme','Homme'], ['sexe','femme','Femme'],
        # education
        ['education','non_scolarise','Non scolarisé'], ['education','primaire','Primaire'],
        ['education','secondaire','Secondaire'], ['education','superieur','Supérieur'],
        # statut_matrimonial
        ['statut_matrimonial','marie','Marié(e)'], ['statut_matrimonial','celibataire','Célibataire'],
        ['statut_matrimonial','concubinage','Concubinage'],
        # saison
        ['saison','saison1','Saison 1'], ['saison','saison2','Saison 2'],
        ['saison','saison3','Saison 3'], ['saison','saison4','Saison 4'], ['saison','autre','Autre'],
        # region (liste officielle 2026)
        ['region','abidjan','Abidjan'], ['region','agno','Agnéby-Tiassa'], ['region','bafing','Bafing'],
        ['region','bagoue','Bagoué'], ['region','belier','Bélier'], ['region','bere','Béré'],
        ['region','bounkani','Bounkani'], ['region','cavally','Cavally'], ['region','folon','Folon'],
        ['region','gbeke','Gbêkê'], ['region','goh','Gôh'], ['region','gontougo','Gontougo'],
        ['region','guinto','Guémon'], ['region','hambol','Hambol'], ['region','hauts_sassandra','Haut-Sassandra'],
        ['region','ifou','Iffou'], ['region','indenié','Indénié-Djuablin'], ['region','kabadougou','Kabadougou'],
        ['region','la_me','La Mé'], ['region','loh_djiboua','Lôh-Djiboua'], ['region','marahoue','Marahoué'],
        ['region','moronou','Moronou'], ['region','nawa','Nawa'], ['region','nzi','N’zi'],
        ['region','pore','Poro'], ['region','san_pedro','San-Pédro'], ['region','sassandra','Sassandra-Marahoué'],
        ['region','savanes','Savanes'], ['region','sud_comoe','Sud-Comoé'], ['region','tonkpi','Tonkpi'],
        ['region','worodougou','Worodougou'], ['region','yamoussoukro','Yamoussoukro'], ['region','zanzan','Zanzan'],
        # type_miel
        ['type_miel','foret','Forêt'], ['type_miel','savane','Savane'], ['type_miel','multifloral','Multifloral'], ['type_miel','autre_miel','Autre'],
        # causes_pertes
        ['causes_pertes','climat','Climat'], ['causes_pertes','mauvaises_pratiques','Mauvaises pratiques apicoles'],
        ['causes_pertes','mauvaise_conservation','Mauvaise conservation'], ['causes_pertes','insectes_animaux','Insectes / animaux'],
        ['causes_pertes','feux_brousse','Feux de brousse'], ['causes_pertes','pesticides','Pesticides'],
        ['causes_pertes','vol_vandalisme','Vol / vandalisme'], ['causes_pertes','autre_cause','Autre'],
        # strategies_reduction
        ['strategies_reduction','amelioration_techniques','Amélioration des techniques apicoles'],
        ['strategies_reduction','equipements','Équipements adaptés'], ['strategies_reduction','bonnes_pratiques','Bonnes pratiques d’hygiène'],
        ['strategies_reduction','stockage','Stockage amélioré'], ['strategies_reduction','protection_ruches','Protection des ruches'],
        ['strategies_reduction','formation','Formation / sensibilisation'], ['strategies_reduction','autre_strategy','Autre'],
        # presence
        ['presence','eau','Eau'], ['presence','ombre','Ombre'], ['presence','aeration','Aération'], ['presence','protection_feux','Protection contre feux'],
        # plantes
        ['plantes','karite','Karité'], ['plantes','anacardier','Anacardier'], ['plantes','manguier','Manguier'], ['plantes','nere','Néré'], ['plantes','eucalyptus','Eucalyptus'], ['plantes','autres','Autres'],
        # zone
        ['zone','foret','Forêt'], ['zone','savane','Savane'], ['zone','agricole','Agricole'], ['zone','mixte','Mixte'],
        # type_ruche
        ['type_ruche','traditionnelle','Traditionnelle'], ['type_ruche','kenyane','Kenyane'], ['type_ruche','dadant','Dadant'], ['type_ruche','autre_ruche','Autre'],
        # equipements
        ['equipements','enfumoir','Enfumoir'], ['equipements','tenue','Tenue'], ['equipements','extracteur','Extracteur'],
        ['equipements','maturateur','Maturateur'], ['equipements','seaux_alimentaires','Seaux alimentaires'],
        # methodes_recolte
        ['methodes_recolte','enfumage','Enfumage'], ['methodes_recolte','pressage','Pressage'], ['methodes_recolte','extraction_centrifuge','Extraction centrifuge'],
        # stockage
        ['stockage','sec','Sec'], ['stockage','abri_lumiere','À l’abri de la lumière'], ['stockage','hermetique','Hermétique'],
        # decantation
        ['decantation','moins_24h','< 24 h'], ['decantation','24_48h','24–48 h'], ['decantation','plus_48h','> 48 h'],
        # potentiel
        ['potentiel','faible','Faible'], ['potentiel','moyen','Moyen'], ['potentiel','eleve','Élevé'],
        # facteur_cle
        ['facteur_cle','financement','Financement'], ['facteur_cle','formation','Formation'], ['facteur_cle','equipements','Équipements'],
        ['facteur_cle','environnement','Environnement'], ['facteur_cle','organisation','Organisation collective'],
        # horizon
        ['horizon','moins_1an','< 1 an'], ['horizon','1_2ans','1–2 ans'], ['horizon','3_5ans','3–5 ans'], ['horizon','5_10ans','5–10 ans'],
        # frequence_visite
        ['frequence_visite','mensuelle','Mensuelle'], ['frequence_visite','bimensuelle','Bimensuelle'], ['frequence_visite','occasionnelle','Occasionnelle'],
        # mode_peuplement
        ['mode_peuplement','essaims_sauvages','Essaims sauvages'], ['mode_peuplement','introduction','Introduction'], ['mode_peuplement','attractifs','Attractifs'],
        # materiel
        ['materiel','inox','Inox'], ['materiel','plastique','Plastique alimentaire'],
        # controle
        ['controle','reine','Reine'], ['controle','couvain','Couvain'], ['controle','reserves','Réserves'], ['controle','maladies','Maladies'],
        # partenaires
        ['partenaires','minader','MINADER'], ['partenaires','anader','ANADER'], ['partenaires','firca','FIRCA'], ['partenaires','ong','ONG'], ['partenaires','financieres','Institutions financières'],
        # appuis
        ['appuis','formation','Formation'], ['appuis','equipements','Équipements'], ['appuis','financement','Financement'], ['appuis','technique','Appui technique'],
        # contraintes
        ['contraintes','climat','Climat'], ['contraintes','materiel','Matériel'], ['contraintes','financement','Financement'], ['contraintes','florales','Ressources florales'],
        # choc_scenario
        ['choc_scenario','-10','-10 %'], ['choc_scenario','-25','-25 %'], ['choc_scenario','-50','-50 %'], ['choc_scenario','plus_50','> -50 %'],
        # amelioration_scenario
        ['amelioration_scenario','+10','+10 %'], ['amelioration_scenario','+25','+25 %'], ['amelioration_scenario','+50','+50 %'], ['amelioration_scenario','plus_50','> +50 %'],
        # refractometre
        ['refractometre','refractometre','Réfractomètre'], ['refractometre','autre','Autre']
    ], columns=['list_name','name','label'])

# Onglet survey (structure complète, modules A à I)
def get_survey():
    survey = []
    add = survey.append
    # Instructions générales
    add(['note','instructions','Objectif général : Collecter des données harmonisées, fiables et exploitables auprès des apiculteurs/producteurs, afin d’analyser la production de miel, la qualité, les contraintes, les appuis reçus et de simuler des scénarios de développement (stress tests) de la filière en Côte d’Ivoire.\n\nInstructions à l’enquêteur :\n- Cocher les cases correspondantes\n- Renseigner les champs numériques ou texte\n- Les sections « par saison » sont à répéter pour chaque saison de récolte observée','','','',''])
    # MODULE A
    add(['note','module_a','MODULE A – IDENTIFICATION ET PROFIL DU PRODUCTEUR','','','',''])
    add(['text','enqueteur_nom','Nom et prénom de l’enquêteur','yes','','',''])
    add(['date','date_entretien','Date de l’entretien','yes','','',''])
    add(['select_one region','region','Région','yes','','',''])
    add(['text','departement','Département','yes','','',''])
    add(['text','localite','Localité / Village','yes','','',''])
    add(['text','activite','Activité (Apiculteur/Producteur de miel)','yes','','',''])
    add(['select_one oui_non','membre_cooperative','Êtes-vous membre d’une organisation ou coopérative apicole ?','yes','','',''])
    add(['text','nom_structure','Si oui, nom de la structure','','','','${membre_cooperative}=\'oui\''])
    add(['integer','age','Âge du producteur (ans)','yes','','',''])
    add(['select_one sexe','sexe','Sexe','yes','','',''])
    add(['text','nationalite','Nationalité','yes','','',''])
    add(['text','telephone','Téléphone','yes','','',''])
    add(['select_one education','education','Niveau d’éducation','yes','','',''])
    add(['select_one statut_matrimonial','statut_matrimonial','Statut matrimonial','yes','','',''])
    add(['integer','taille_menage','Taille totale du ménage','yes','','',''])
    add(['integer','enfants','Enfants (0–14 ans)','yes','','',''])
    add(['integer','jeunes','Jeunes (15–35 ans)','yes','','',''])
    add(['integer','adultes','Adultes (36–59 ans)','yes','','',''])
    add(['integer','personnes_agees','Personnes âgées (≥ 60 ans)','yes','','',''])
    add(['integer','main_oeuvre','Nombre total de personnes impliquées dans l’activité apicole','yes','','',''])
    add(['text','repartition_age','Répartition par âge (si possible)','no','','',''])
    # MODULE B
    add(['note','module_b','MODULE B – EXPÉRIENCE ET ORGANISATION DE LA PRODUCTION','','','',''])
    add(['integer','annees_experience','Années d’expérience en apiculture','yes','','',''])
    add(['integer','nb_saisons','Nombre de saisons de récolte par an','yes','','',''])
    add(['begin_repeat','saisons','Informations par saison','','','',''])
    add(['select_one saison','saison','Saison concernée','yes','','',''])
    add(['integer','volume_produit','Volume total produit (kg)','yes','','',''])
    add(['select_multiple type_miel','type_miel','Type(s) de miel','yes','','',''])
    add(['integer','nombre_ruches','Nombre total de ruches installées','yes','','',''])
    add(['integer','ruches_productives','Nombre de ruches productives','yes','','',''])
    add(['integer','production_par_ruche','Production moyenne par ruche (kg/ruche)','yes','','',''])
    add(['integer','autoconsommation','Quantité autoconsommée (kg)','yes','','',''])
    add(['integer','vendue','Quantité vendue (kg)','yes','','',''])
    add(['decimal','prix_vente','Prix moyen de vente au producteur (FCFA/kg)','yes','','',''])
    add(['end_repeat','saisons','','','','',''])
    # MODULE C
    add(['note','module_c','MODULE C – PERTES DE PRODUCTION','','','',''])
    add(['select_one oui_non','pertes','Avez-vous subi des pertes de miel ou de colonies ?','yes','','',''])
    add(['integer','volume_perdu','Si oui, volume estimé perdu (kg)','','','','${pertes}=\'oui\''])
    add(['select_multiple causes_pertes','causes_pertes','Causes des pertes','no','','','${pertes}=\'oui\''])
    add(['select_multiple strategies_reduction','strategies_reduction','Stratégies de réduction des pertes','no','','','${pertes}=\'oui\''])
    # MODULE D
    add(['note','module_d','MODULE D – CONDITIONS DE PRODUCTION ET QUALITÉ','','','',''])
    add(['select_one zone','zone','Zone','yes','','',''])
    add(['select_multiple presence','presence','Présence (Eau, Ombre, Aération, Protection contre feux)','no','','',''])
    add(['select_multiple plantes','plantes','Plantes mellifères dominantes','no','','',''])
    add(['select_one oui_non','pesticides','Exposition aux pesticides','yes','','',''])
    add(['select_multiple type_ruche','type_ruche','Types de ruches','yes','','',''])
    add(['select_multiple equipements','equipements','Équipements disponibles','no','','',''])
    add(['select_one mode_peuplement','mode_peuplement','Mode de peuplement','no','','',''])
    add(['select_one frequence_visite','frequence_visite','Fréquence de visite','no','','',''])
    add(['select_multiple controle','controle','Contrôles réalisés','no','','',''])
    add(['select_one oui_non','alveoles_operculees','Alvéoles operculées à la récolte','no','','',''])
    add(['select_one methodes_recolte','methodes_recolte','Méthode de récolte','no','','',''])
    add(['select_one oui_non','couvain_evite','Couvain évité','no','','',''])
    add(['select_one oui_non','filtration','Filtration','no','','',''])
    add(['select_one decantation','decantation','Décantation','no','','',''])
    add(['select_multiple stockage','stockage','Stockage','no','','',''])
    # MODULE E
    add(['note','module_e','MODULE E – APPUIS ET STRUCTURATION','','','',''])
    add(['select_multiple appuis','appuis','Appuis reçus','no','','',''])
    add(['select_multiple partenaires','partenaires','Partenaires','no','','',''])
    add(['select_multiple facteur_cle','besoins_prioritaires','Besoins prioritaires','no','','',''])
    # MODULE F
    add(['note','module_f','MODULE F – CONFORMITÉ ET QUALITÉ','','','',''])
    add(['select_multiple refractometre','controle_maturite','Contrôle de maturité','no','','',''])
    add(['integer','temperature_extraction','Température maximale d’extraction (°C)','no','','',''])
    add(['select_multiple materiel','materiel','Matériel utilisé','no','','',''])
    add(['select_one oui_non','analyses_labo','Analyses en laboratoire','no','','',''])
    add(['text','analyses_preciser','Si oui, préciser','','','','${analyses_labo}=\'oui\''])
    # MODULE G
    add(['note','module_g','MODULE G – PERSPECTIVES ET POTENTIEL','','','',''])
    add(['select_one potentiel','potentiel','Potentiel apicole de la zone','no','','',''])
    add(['select_multiple contraintes','contraintes','Contraintes majeures','no','','',''])
    add(['integer','prod_estimee_si_levier','Production annuelle estimée si contraintes levées (kg/an)','no','','',''])
    # MODULE H
    add(['note','module_h','MODULE H – STRESS TEST ET SCÉNARIOS DE DÉVELOPPEMENT','','','',''])
    add(['integer','production_actuelle','Production annuelle actuelle (kg/an)','yes','','',''])
    add(['integer','ruches_productives_tot','Nombre de ruches productives','yes','','',''])
    add(['decimal','rendement_moyen','Rendement moyen par ruche (kg/ruche/an)','yes','','',''])
    add(['select_multiple amelioration_scenario','amelioration_equipements','Amélioration des équipements','no','','',''])
    add(['select_multiple amelioration_scenario','formation_continue','Formation technique continue','no','','',''])
    add(['integer','ruches_supplementaires','Ruches supplémentaires possibles','no','','',''])
    add(['integer','prod_apres_augmentation','Production estimée après augmentation (kg/an)','no','','',''])
    add(['integer','pertes_actuelles','Pertes actuelles (%)','no','','',''])
    add(['integer','pertes_apres_amelioration','Pertes après amélioration (%)','no','','',''])
    add(['select_multiple choc_scenario','secheresse','Sécheresse prolongée','no','','',''])
    add(['select_multiple choc_scenario','feux_brousse','Feux de brousse','no','','',''])
    add(['select_multiple choc_scenario','pesticides','Usage accru de pesticides','no','','',''])
    add(['select_multiple choc_scenario','maladies','Maladies / mortalité des abeilles','no','','',''])
    add(['select_one facteur_cle','facteur_impact','Quel facteur unique aurait le plus grand impact ?','no','','',''])
    add(['select_one horizon','horizon','Délai pour observer une augmentation significative','no','','',''])
    # MODULE I
    add(['note','module_i','MODULE I – OBSERVATIONS','','','',''])
    add(['text','commentaires','Commentaires du producteur','no','','',''])
    add(['text','observations','Observations de l’enquêteur','no','','',''])
    return pd.DataFrame(survey, columns=['type','name','label','required','appearance','repeat_count','relevant'])

# Export Excel
settings = get_settings()
survey = get_survey()
choices = get_choices()

with pd.ExcelWriter('c:/Users/DELL/OneDrive - BIZAO/Pictures/solve/questionnaire_miel_kobo_complet_A_I.xlsx') as writer:
    settings.to_excel(writer, sheet_name='settings', index=False)
    survey.to_excel(writer, sheet_name='survey', index=False)
    choices.to_excel(writer, sheet_name='choices', index=False)
