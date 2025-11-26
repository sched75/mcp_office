
# 📚 Référence API - MCP Office

## Introduction

Cette documentation détaille les **271 outils** disponibles dans le serveur MCP Office. Chaque outil est accessible via Claude Desktop avec le format `application_nom_methode`.

---

## 📝 Word (59 outils)

### Gestion Documents
- **`word_create_document`** - Crée un nouveau document Word
- **`word_open_document`** - Ouvre un document existant
- **`word_save_document`** - Enregistre le document
- **`word_close_document`** - Ferme le document
- **`word_export_to_pdf`** - Exporte en PDF
- **`word_print_to_pdf`** - Imprime en PDF (alias)

### Modèles
- **`word_create_from_template`** - Crée depuis un modèle
- **`word_save_as_template`** - Sauvegarde comme modèle
- **`word_list_available_templates`** - Liste les modèles disponibles

### Contenu Textuel
- **`word_add_paragraph`** - Ajoute un paragraphe
- **`word_insert_text_at_position`** - Insère du texte à une position
- **`word_find_and_replace`** - Recherche et remplace
- **`word_delete_text`** - Supprime du texte

### Formatage
- **`word_apply_text_formatting`** - Applique le formatage
- **`word_set_paragraph_alignment`** - Définit l'alignement
- **`word_apply_style`** - Applique un style prédéfini
- **`word_set_line_spacing`** - Définit l'interligne
- **`word_create_custom_style`** - Crée un style personnalisé

### Tableaux
- **`word_insert_table`** - Insère un tableau
- **`word_set_table_cell_text`** - Remplit une cellule
- **`word_add_table_row`** - Ajoute une ligne
- **`word_add_table_column`** - Ajoute une colonne
- **`word_delete_table_row`** - Supprime une ligne
- **`word_delete_table_column`** - Supprime une colonne
- **`word_merge_table_cells`** - Fusionne des cellules

### Images et Objets
- **`word_insert_image`** - Insère une image
- **`word_insert_image_from_clipboard`** - Insère depuis presse-papiers
- **`word_resize_image`** - Redimensionne l'image
- **`word_position_image`** - Positionne l'image
- **`word_crop_image`** - Recadre l'image
- **`word_apply_image_effects`** - Applique des effets
- **`word_insert_shape`** - Insère une forme
- **`word_add_textbox`** - Ajoute une zone de texte

### Structure Document
- **`word_add_header`** - Ajoute un en-tête
- **`word_add_footer`** - Ajoute un pied de page
- **`word_insert_page_numbers`** - Insère des numéros de page
- **`word_create_table_of_contents`** - Crée une table des matières
- **`word_insert_page_break`** - Insère un saut de page
- **`word_insert_section_break`** - Insère un saut de section
- **`word_configure_section`** - Configure une section

### Révision
- **`word_enable_track_changes`** - Active le suivi des modifications
- **`word_disable_track_changes`** - Désactive le suivi
- **`word_add_comment`** - Ajoute un commentaire
- **`word_accept_all_revisions`** - Accepte toutes les révisions
- **`word_reject_all_revisions`** - Rejette toutes les révisions

### Métadonnées
- **`word_get_document_properties`** - Obtient les propriétés
- **`word_set_document_properties`** - Définit les propriétés
- **`word_get_document_statistics`** - Obtient les statistiques
- **`word_set_document_language`** - Définit la langue

### Impression
- **`word_configure_print_settings`** - Configure l'impression
- **`word_print_preview`** - Aperçu avant impression

### Protection
- **`word_protect_document`** - Protège le document
- **`word_set_password`** - Définit un mot de passe
- **`word_unprotect_document`** - Supprime la protection

### Fonctionnalités Avancées
- **`word_mail_merge_with_data`** - Publipostage
- **`word_insert_bookmark`** - Insère un signet
- **`word_create_index`** - Crée un index
- **`word_manage_bibliography`** - Gère la bibliographie
- **`word_insert_field`** - Insère un champ
- **`word_compare_documents`** - Compare des documents
- **`word_insert_smartart`** - Insère SmartArt
- **`word_convert_format`** - Convertit le format
- **`word_modify_style`** - Modifie un style
- **`word_insert_hyperlink`** - Insère un lien hypertexte

---

## 📊 Excel (82 outils)

### Gestion Classeurs
- **`excel_create_workbook`** - Crée un classeur
- **`excel_open_workbook`** - Ouvre un classeur
- **`excel_save_workbook`** - Sauvegarde le classeur
- **`excel_close_workbook`** - Ferme le classeur
- **`excel_export_to_pdf`** - Exporte en PDF
- **`excel_convert_to_csv`** - Convertit en CSV

### Modèles
- **`excel_create_from_template`** - Crée depuis modèle
- **`excel_save_as_template`** - Sauvegarde comme modèle
- **`excel_list_custom_templates`** - Liste les modèles

### Gestion Feuilles
- **`excel_add_worksheet`** - Ajoute une feuille
- **`excel_delete_worksheet`** - Supprime une feuille
- **`excel_rename_worksheet`** - Renomme une feuille
- **`excel_copy_worksheet`** - Copie une feuille
- **`excel_move_worksheet`** - Déplace une feuille
- **`excel_hide_worksheet`** - Masque une feuille
- **`excel_show_worksheet`** - Affiche une feuille

### Cellules et Données
- **`excel_write_cell`** - Écrit dans une cellule
- **`excel_write_range`** - Écrit dans une plage
- **`excel_read_cell`** - Lit une cellule
- **`excel_read_range`** - Lit une plage
- **`excel_copy_paste_cells`** - Copie-colle des cellules
- **`excel_clear_contents`** - Efface le contenu
- **`excel_find_and_replace`** - Recherche et remplace

### Formules et Calculs
- **`excel_write_formula`** - Écrit une formule
- **`excel_use_function`** - Utilise une fonction
- **`excel_use_vlookup`** - Utilise VLOOKUP
- **`excel_set_reference_type`** - Définit le type de référence
- **`excel_use_array_formula`** - Applique une formule matricielle

### Formatage
- **`excel_set_number_format`** - Définit le format numérique
- **`excel_set_cell_color`** - Définit la couleur de cellule
- **`excel_set_font_color`** - Définit la couleur de police
- **`excel_set_borders`** - Définit les bordures
- **`excel_set_alignment`** - Définit l'alignement
- **`excel_set_wrap_text`** - Définit le retour à la ligne
- **`excel_merge_cells`** - Fusionne des cellules
- **`excel_set_column_width`** - Définit la largeur de colonne
- **`excel_set_row_height`** - Définit la hauteur de ligne
- **`excel_conditional_formatting`** - Applique le formatage conditionnel

### Tableaux Structurés
- **`excel_convert_to_table`** - Convertit en tableau
- **`excel_add_total_row`** - Ajoute une ligne de total
- **`excel_apply_table_style`** - Applique un style de tableau
- **`excel_filter_table`** - Filtre le tableau
- **`excel_sort_table`** - Trie le tableau

### Images et Objets
- **`excel_insert_image`** - Insère une image
- **`excel_resize_image`** - Redimensionne l'image
- **`excel_position_image`** - Positionne l'image
- **`excel_anchor_image_to_cell`** - Ancre l'image à une cellule
- **`excel_insert_logo_watermark`** - Insère un filigrane

### Graphiques
- **`excel_create_chart`** - Crée un graphique
- **`excel_modify_chart_data`** - Modifie les données du graphique
- **`excel_customize_chart_title`** - Personnalise le titre
- **`excel_customize_chart_legend`** - Personnalise la légende
- **`excel_modify_chart_axes`** - Modifie les axes
- **`excel_change_chart_colors`** - Change les couleurs
- **`excel_move_resize_chart`** - Déplace et redimensionne

### Tableaux Croisés Dynamiques
- **`excel_create_pivot_table`** - Crée un tableau croisé
- **`excel_set_pivot_fields`** - Définit les champs
- **`excel_apply_pivot_filter`** - Applique un filtre
- **`excel_change_pivot_calculation`** - Change le calcul
- **`excel_refresh_pivot_table`** - Actualise le tableau

### Tri et Filtres
- **`excel_sort_ascending`** - Trie ascendant
- **`excel_sort_descending`** - Trie descendant
- **`excel_apply_autofilter`** - Applique l'auto-filtre
- **`excel_create_advanced_filter`** - Crée un filtre avancé

### Protection
- **`excel_protect_worksheet`** - Protège la feuille
- **`excel_protect_workbook`** - Protège le classeur
- **`excel_set_workbook_password`** - Définit un mot de passe
- **`excel_unprotect_worksheet`** - Supprime la protection

### Plages Nommées
- **`excel_create_named_range`** - Crée une plage nommée
- **`excel_use_named_range_in_formula`** - Utilise une plage nommée
- **`excel_delete_named_range`** - Supprime une plage nommée

### Validation de Données
- **`excel_create_dropdown_list`** - Crée une liste déroulante
- **`excel_set_validation_rules`** - Définit les règles de validation
- **`excel_remove_validation`** - Supprime la validation

### Impression
- **`excel_configure_print_settings`** - Configure l'impression
- **`excel_set_print_area`** - Définit la zone d'impression
- **`excel_print_preview`** - Aperçu avant impression

### Fonctionnalités Avancées
- **`excel_group_rows_columns`** - Groupe lignes/colonnes
- **`excel_freeze_panes`** - Figer les volets
- **`excel_split_window`** - Diviser la fenêtre
- **`excel_create_sparklines`** - Crée des sparklines
- **`excel_scenario_analysis`** - Analyse de scénarios
- **`excel_goal_seek`** - Valeur cible
- **`excel_use_solver`** - Utilise le solveur
- **`excel_consolidate_data`** - Consolide les données
- **`excel_create_subtotals`** - Crée des sous-totaux
- **`excel_import_csv`** - Importe CSV
- **`excel_insert_hyperlink`** - Insère un lien hypertexte
- **`excel_insert_comment`** - Insère un commentaire
- **`excel_use_3d_reference`** - Utilise une référence 3D
- **`excel_export_to_json`** - Exporte en JSON

---

## 🎨 PowerPoint (63 outils)

### Gestion Présentations
- **`powerpoint_create_presentation`** - Crée une présentation
- **`powerpoint_open_presentation`** - Ouvre une présentation
- **`powerpoint_save_presentation`** - Sauvegarde la présentation
- **`powerpoint_close_presentation`** - Ferme la présentation
- **`powerpoint_export_to_pdf`** - Exporte en PDF
- **`powerpoint_save_as`** - Sauvegarde sous
- **`powerpoint_create_from_template`** - Crée depuis modèle
- **`powerpoint_save_as_template`** - Sauvegarde comme modèle
- **`powerpoint_apply_template`** - Applique un modèle
- **`powerpoint_create_custom_slide_master`** - Crée un masque personnalisé

### Gestion Diapositives
- **`powerpoint_add_slide`** - Ajoute une diapositive
- **`powerpoint_delete_slide`** - Supprime une diapositive
- **`powerpoint_duplicate_slide`** - Duplique une diapositive
- **`powerpoint_move_slide`** - Déplace une diapositive
- **`powerpoint_apply_slide_layout`** - Applique un layout
- **`powerpoint_hide_show_slide`** - Masque/affiche une diapositive

### Contenu Textuel
- **`powerpoint_add_textbox`** - Ajoute une zone de texte
- **`powerpoint_modify_title`** - Modifie le titre
- **`powerpoint_modify_body_text`** - Modifie le texte principal
- **`powerpoint_add_bullets`** - Ajoute des puces
- **`powerpoint_add_numbered_list`** - Ajoute une liste numérotée
- **`powerpoint_format_text`** - Formate le texte

### Images et Médias
- **`powerpoint_insert_image`** - Insère une image
- **`powerpoint_resize_image`** - Redimensionne l'image
- **`powerpoint_reposition_image`** - Repositionne l'image
- **`powerpoint_insert_video`** - Insère une vidéo
- **`powerpoint_insert_audio`** - Insère un audio

### Formes et Objets
- **`powerpoint_insert_shape`** - Insère une forme
- **`powerpoint_modify_fill_color`** - Modifie la couleur de remplissage
- **`powerpoint_modify_outline`** - Modifie le contour
- **`powerpoint_group_shapes`** - Groupe des formes
- **`powerpoint_ungroup_shapes`** - Dégroupe des formes

### Tableaux
- **`powerpoint_insert_table`** - Insère un tableau
- **`powerpoint_fill_table_cell`** - Remplit une cellule
- **`powerpoint_merge_table_cells`** - Fusionne des cellules
- **`powerpoint_split_table_cell`** - Divise une cellule
- **`powerpoint_apply_table_style`** - Applique un style de tableau
- **`powerpoint_format_table_borders`** - Formate les bordures

### Graphiques
- **`powerpoint_insert_chart`** - Insère un graphique
- **`powerpoint_link_excel_chart`** - Lie un graphique Excel
- **`powerpoint_modify_chart_data`** - Modifie les données
- **`powerpoint_customize_chart_style`** - Personnalise le style

### Animations
- **`powerpoint_add_entrance_animation`** - Ajoute une animation d'entrée
- **`powerpoint_add_exit_animation`** - Ajoute une animation de sortie
- **`powerpoint_set_animation_order`** - Définit l'ordre des animations
- **`powerpoint_configure_animation_timing`** - Configure le timing

### Transitions
- **`powerpoint_apply_transition`** - Applique une transition
- **`powerpoint_set_transition_duration`** - Définit la durée
- **`powerpoint_apply_transition_to_all`** - Applique à toutes

### Thèmes et Design
- **`powerpoint_apply_theme`** - Applique un thème
- **`powerpoint_modify_color_scheme`** - Modifie le schéma de couleurs
- **`powerpoint_modify_theme_fonts`** - Modifie les polices du thème
- **`powerpoint_set_background`** - Définit l'arrière-plan
- **`powerpoint_apply_slide_master`** - Applique un masque de diapositive

### Notes et Commentaires
- **`powerpoint_add_speaker_notes`** - Ajoute des notes d'orateur
- **`powerpoint_read_speaker_notes`** - Lit les notes d'orateur
- **`powerpoint_add_comment`** - Ajoute un commentaire

### Fonctionnalités Avancées
- **`powerpoint_start_presenter_mode`** - Démarre le mode présentateur
- **`powerpoint_set_slide_timing`** - Définit le timing des diapositives
- **`powerpoint_record_slideshow`** - Enregistre le diaporama
- **`powerpoint_insert_smartart`** - Insère SmartArt
- **`powerpoint_insert_ole_object`** - Insère un objet OLE
- **`powerpoint_create_section_zoom`** - Crée un zoom de section
- **`powerpoint_insert_hyperlink`** - Insère un lien hypertexte
- **`powerpoint_add_action_trigger`** - Ajoute un déclencheur d'action
- **`powerpoint_export_to_video`** - Exporte en vidéo
- **`powerpoint_add_captions`** - Ajoute des légendes
- **`powerpoint_compare_presentations`** - Compare des présentations

---

## 📧 Outlook (67 outils)

### Emails
- **`outlook_send_email`** - Envoie un email
- **`outlook_send_with_attachments`** - Envoie avec pièces jointes
- **`outlook_read_email`** - Lit un email
- **`outlook_reply_to_email`** - Répond à un email
- **`outlook_reply_all_to_email`** - Répond à tous
- **`outlook_forward_email`** - Transfère un email
- **`outlook_mark_as_read`** - Marque comme lu
- **`outlook_mark_as_unread`** - Marque comme non lu
- **`outlook_flag_email`** - Ajoute un drapeau
- **`outlook_delete_email`** - Supprime un email
- **`outlook_move_email_to_folder`** - Déplace vers un dossier
- **`outlook_search_emails`** - Recherche des emails

### Pièces Jointes
- **`outlook_add_attachment`** - Ajoute une pièce jointe
- **`outlook_list_attachments`** - Liste les pièces jointes
- **`outlook_save_attachment`** - Sauvegarde une pièce jointe
- **`outlook_remove_attachment`** - Supprime une pièce jointe
- **`outlook_create_new_message`** - Crée un nouveau brouillon

### Dossiers
- **`outlook_create_folder`** - Crée un dossier
- **`outlook_delete_folder`** - Supprime un dossier
- **`outlook_rename_folder`** - Renomme un dossier
- **`outlook_move_folder`** - Déplace un dossier
- **`outlook_list_folders`** - Liste les dossiers
- **`outlook_get_folder_item_count`** - Compte les éléments d'un dossier
- **`outlook_get_unread_count`** - Compte les messages non lus

### Calendrier
- **`outlook_create_appointment`** - Crée un rendez-vous
- **`outlook_create_recurring_event`** - Crée un événement récurrent
- **`outlook_read_appointment`** - Lit un rendez-vous
- **`outlook_modify_appointment`** - Modifie un rendez-vous
- **`outlook_delete_appointment`** - Supprime un rendez-vous
- **`outlook_search_appointments`** - Recherche des rendez-vous
- **`outlook_get_appointments_by_date`** - Obtient par date
- **`outlook_set_reminder`** - Définit un rappel
- **`outlook_set_busy_status`** - Définit le statut occupé
- **`outlook_export_appointment_ics`** - Exporte en ICS
- **`outlook_get_calendar_count`** - Compte les rendez-vous
- **`outlook_export_to_pdf`** - Exporte en PDF

### Réunions
- **`outlook_create_meeting_request`** - Crée une demande de réunion
- **`outlook_invite_participants`** - Invite des participants
- **`outlook_accept_meeting`** - Accepte une réunion
- **`outlook_decline_meeting`** - Refuse une réunion
- **`outlook_propose_new_time`** - Propose un nouveau créneau
- **`outlook_cancel_meeting`** - Annule une réunion
- **`outlook_update_meeting`** - Met à jour une réunion
- **`outlook_check_availability`** - Vérifie la disponibilité

### Contacts
- **`outlook_create_contact`** - Crée un contact
- **`outlook_modify_contact`** - Modifie un contact
- **`outlook_delete_contact`** - Supprime un contact
- **`outlook_search_contact`** - Recherche un contact
- **`outlook_list_all_contacts`** - Liste tous les contacts
- **`outlook_create_contact_group`** - Crée un groupe de contacts
- **`outlook_add_to_contact_group`** - Ajoute à un groupe
- **`outlook_export_contacts_vcf`** - Exporte en VCF
- **`outlook_import_contacts`** - Importe des contacts

### Tâches
- **`outlook_create_task`** - Crée une tâche
- **`outlook_modify_task`** - Modifie une tâche
- **`outlook_delete_task`** - Supprime une tâche
- **`outlook_mark_task_complete`** - Marque comme terminée
- **`outlook_set_task_priority`** - Définit la priorité
- **`outlook_set_task_due_date`** - Définit l'échéance
- **`outlook_list_tasks`** - Liste les tâches

### Utilitaires
- **`outlook_list_accounts`** - Liste les comptes
- **`outlook_get_default_account`** - Obtient le compte par défaut
- **`outlook_get_inbox_count`** - Compte les messages inbox
- **`outlook_create_category`** - Crée une catégorie
- **`outlook_list_categories`** - Liste les catégories
- **`outlook_apply_category`** - Applique une catégorie
- **`outlook_com_operation`** - Opération COM personnalisée

---

## 🔧 Utilisation des Outils

### Format des Commandes
Tous les outils suivent le format : `application_nom_methode`

**Exemples :**
- `word_create_document`
- `excel_write_cell`
- `powerpoint_add_slide`
- `outlook_send_email`

### Paramètres Requis
Chaque outil a des paramètres spécifiques. Consultez la configuration dans `src/tools_configs.py` pour les détails complets.

### Gestion des Erreurs
- ✅ Retourne `success: true` en cas de succès
- ❌ Retourne `success: false` avec `error` en cas d'échec
- 🔧 Gestion robuste des exceptions COM

### Performance
- ⚡ Initialisation rapide des services
- 🔄 Gestion automatique des connexions COM
- 🧹 Nettoyage automatique des ressources

---

## 📞 Support

Pour toute question sur l'utilisation des outils :
- 📖 Consultez le [Guide Utilisateur](user_guide.md)
- 🔧 Voir le [Troubleshooting](troubleshooting.md)
- 🐛 [Issues GitHub](https://github.com/sched75/mcp_office/issues)

**Profitez de l'automation complète d'Office avec Claude ! 🚀**
