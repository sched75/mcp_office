"""
Configurations des outils MCP pour Word, Excel et PowerPoint.
Généré automatiquement.
"""

WORD_TOOLS_CONFIG = {
    "export_to_pdf": {
        "required": [
            "output_path"
        ],
        "optional": [],
        "desc": "Export document to PDF."
    },
    "print_to_pdf": {
        "required": [
            "output_path"
        ],
        "optional": [],
        "desc": "Print document to PDF (alias for export_to_pdf)."
    },
    "create_from_template": {
        "required": [
            "template_path"
        ],
        "optional": [],
        "desc": "Create document from template."
    },
    "save_as_template": {
        "required": [
            "template_path"
        ],
        "optional": [],
        "desc": "Save current document as template."
    },
    "list_available_templates": {
        "required": [],
        "optional": [
            "directory"
        ],
        "desc": "List available Word templates."
    },
    "add_paragraph": {
        "required": [
            "text"
        ],
        "optional": [
            "style"
        ],
        "desc": "Add a paragraph to the document."
    },
    "insert_text_at_position": {
        "required": [
            "text"
        ],
        "optional": [
            "position"
        ],
        "desc": "Insert text at specific position."
    },
    "find_and_replace": {
        "required": [
            "find_text",
            "replace_text"
        ],
        "optional": [
            "match_case"
        ],
        "desc": "Find and replace text in document."
    },
    "delete_text": {
        "required": [
            "start",
            "end"
        ],
        "optional": [],
        "desc": "Delete text between positions."
    },
    "apply_text_formatting": {
        "required": [],
        "optional": [
            "bold",
            "italic",
            "underline",
            "font_name",
            "font_size",
            "color_rgb",
            "start",
            "end"
        ],
        "desc": "Apply text formatting."
    },
    "set_paragraph_alignment": {
        "required": [
            "alignment"
        ],
        "optional": [
            "paragraph_index"
        ],
        "desc": "Set paragraph alignment."
    },
    "apply_style": {
        "required": [
            "style_name"
        ],
        "optional": [
            "paragraph_index"
        ],
        "desc": "Apply predefined style."
    },
    "set_line_spacing": {
        "required": [
            "spacing"
        ],
        "optional": [
            "paragraph_index"
        ],
        "desc": "Set line spacing."
    },
    "create_custom_style": {
        "required": [
            "style_name"
        ],
        "optional": [
            "base_style"
        ],
        "desc": "Create custom style."
    },
    "insert_table": {
        "required": [
            "rows",
            "cols"
        ],
        "optional": [],
        "desc": "Insert table with dimensions."
    },
    "set_table_cell_text": {
        "required": [
            "table_index",
            "row",
            "col",
            "text"
        ],
        "optional": [],
        "desc": "Set text in table cell."
    },
    "add_table_row": {
        "required": [
            "table_index"
        ],
        "optional": [],
        "desc": "Add row to table."
    },
    "add_table_column": {
        "required": [
            "table_index"
        ],
        "optional": [],
        "desc": "Add column to table."
    },
    "delete_table_row": {
        "required": [
            "table_index",
            "row_index"
        ],
        "optional": [],
        "desc": "Delete row from table."
    },
    "delete_table_column": {
        "required": [
            "table_index",
            "col_index"
        ],
        "optional": [],
        "desc": "Delete column from table."
    },
    "merge_table_cells": {
        "required": [
            "table_index",
            "start_row",
            "start_col",
            "end_row",
            "end_col"
        ],
        "optional": [],
        "desc": "Merge table cells."
    },
    "insert_image": {
        "required": [
            "image_path"
        ],
        "optional": [
            "width",
            "height"
        ],
        "desc": "Insert image from file."
    },
    "insert_image_from_clipboard": {
        "required": [],
        "optional": [],
        "desc": "Insert image from clipboard."
    },
    "resize_image": {
        "required": [
            "image_index",
            "width",
            "height"
        ],
        "optional": [],
        "desc": "Resize image."
    },
    "position_image": {
        "required": [
            "image_index"
        ],
        "optional": [
            "wrap_format"
        ],
        "desc": "Position image with text wrapping."
    },
    "crop_image": {
        "required": [
            "image_index"
        ],
        "optional": [
            "left",
            "top",
            "right",
            "bottom"
        ],
        "desc": "Crop image."
    },
    "apply_image_effects": {
        "required": [
            "image_index"
        ],
        "optional": [
            "brightness",
            "contrast"
        ],
        "desc": "Apply effects to image."
    },
    "insert_shape": {
        "required": [
            "shape_type",
            "left",
            "top",
            "width",
            "height"
        ],
        "optional": [],
        "desc": "Insert shape."
    },
    "add_textbox": {
        "required": [
            "text",
            "left",
            "top",
            "width",
            "height"
        ],
        "optional": [],
        "desc": "Add textbox."
    },
    "add_header": {
        "required": [
            "text"
        ],
        "optional": [
            "section_index"
        ],
        "desc": "Add header to document."
    },
    "add_footer": {
        "required": [
            "text"
        ],
        "optional": [
            "section_index"
        ],
        "desc": "Add footer to document."
    },
    "insert_page_numbers": {
        "required": [],
        "optional": [
            "position"
        ],
        "desc": "Insert page numbers."
    },
    "create_table_of_contents": {
        "required": [],
        "optional": [],
        "desc": "Create table of contents."
    },
    "insert_page_break": {
        "required": [],
        "optional": [],
        "desc": "Insert page break."
    },
    "insert_section_break": {
        "required": [],
        "optional": [
            "break_type"
        ],
        "desc": "Insert section break."
    },
    "configure_section": {
        "required": [
            "section_index"
        ],
        "optional": [
            "orientation",
            "page_width",
            "page_height"
        ],
        "desc": "Configure section properties."
    },
    "enable_track_changes": {
        "required": [],
        "optional": [],
        "desc": "Enable track changes."
    },
    "disable_track_changes": {
        "required": [],
        "optional": [],
        "desc": "Disable track changes."
    },
    "add_comment": {
        "required": [
            "text"
        ],
        "optional": [
            "range_start",
            "range_end"
        ],
        "desc": "Add comment to document."
    },
    "accept_all_revisions": {
        "required": [],
        "optional": [],
        "desc": "Accept all revisions."
    },
    "reject_all_revisions": {
        "required": [],
        "optional": [],
        "desc": "Reject all revisions."
    },
    "get_document_properties": {
        "required": [],
        "optional": [],
        "desc": "Get document properties."
    },
    "set_document_properties": {
        "required": [],
        "optional": [
            "author",
            "title",
            "subject",
            "keywords"
        ],
        "desc": "Set document properties."
    },
    "get_document_statistics": {
        "required": [],
        "optional": [],
        "desc": "Get document statistics."
    },
    "set_document_language": {
        "required": [
            "language_id"
        ],
        "optional": [],
        "desc": "Set document language."
    },
    "configure_print_settings": {
        "required": [],
        "optional": [
            "copies",
            "page_range",
            "collate"
        ],
        "desc": "Configure print settings."
    },
    "print_preview": {
        "required": [],
        "optional": [],
        "desc": "Show print preview."
    },
    "protect_document": {
        "required": [
            "protection_type"
        ],
        "optional": [
            "password"
        ],
        "desc": "Protect document."
    },
    "set_password": {
        "required": [
            "password"
        ],
        "optional": [],
        "desc": "Set document password."
    },
    "unprotect_document": {
        "required": [],
        "optional": [
            "password"
        ],
        "desc": "Remove document protection."
    },
    "mail_merge_with_data": {
        "required": [
            "data_source"
        ],
        "optional": [],
        "desc": "Perform mail merge."
    },
    "insert_bookmark": {
        "required": [
            "name"
        ],
        "optional": [
            "range_start",
            "range_end"
        ],
        "desc": "Insert bookmark."
    },
    "create_index": {
        "required": [],
        "optional": [],
        "desc": "Create index."
    },
    "manage_bibliography": {
        "required": [],
        "optional": [
            "source_file"
        ],
        "desc": "Manage bibliography."
    },
    "insert_field": {
        "required": [
            "field_type"
        ],
        "optional": [
            "text"
        ],
        "desc": "Insert field."
    },
    "compare_documents": {
        "required": [
            "original_path",
            "revised_path"
        ],
        "optional": [],
        "desc": "Compare two documents."
    },
    "insert_smartart": {
        "required": [],
        "optional": [
            "layout"
        ],
        "desc": "Insert SmartArt."
    },
    "convert_format": {
        "required": [
            "output_path",
            "file_format"
        ],
        "optional": [],
        "desc": "Convert document format."
    },
    "modify_style": {
        "required": [
            "style_name"
        ],
        "optional": [],
        "desc": "Modify existing style."
    },
    "insert_hyperlink": {
        "required": [
            "text",
            "url"
        ],
        "optional": [
            "range_start",
            "range_end"
        ],
        "desc": "Insert hyperlink."
    }
}

EXCEL_TOOLS_CONFIG = {
    "create_workbook": {
        "required": [],
        "optional": [],
        "desc": "Create a new workbook."
    },
    "open_workbook": {
        "required": [
            "file_path"
        ],
        "optional": [],
        "desc": "Open an existing workbook."
    },
    "save_workbook": {
        "required": [],
        "optional": [
            "file_path"
        ],
        "desc": "Save the current workbook."
    },
    "close_workbook": {
        "required": [],
        "optional": [
            "save_changes"
        ],
        "desc": "Close the current workbook."
    },
    "export_to_pdf": {
        "required": [
            "output_path"
        ],
        "optional": [],
        "desc": "Export workbook to PDF."
    },
    "convert_to_csv": {
        "required": [
            "output_path"
        ],
        "optional": [],
        "desc": "Convert workbook to CSV."
    },
    "create_from_template": {
        "required": [
            "template_path"
        ],
        "optional": [],
        "desc": "Create workbook from template."
    },
    "save_as_template": {
        "required": [
            "template_path"
        ],
        "optional": [],
        "desc": "Save workbook as template."
    },
    "list_custom_templates": {
        "required": [],
        "optional": [
            "directory"
        ],
        "desc": "List available custom templates."
    },
    "add_worksheet": {
        "required": [],
        "optional": [
            "name"
        ],
        "desc": "Add a new worksheet."
    },
    "delete_worksheet": {
        "required": [
            "sheet_name"
        ],
        "optional": [],
        "desc": "Delete a worksheet."
    },
    "rename_worksheet": {
        "required": [
            "old_name",
            "new_name"
        ],
        "optional": [],
        "desc": "Rename a worksheet."
    },
    "copy_worksheet": {
        "required": [
            "sheet_name"
        ],
        "optional": [
            "new_name"
        ],
        "desc": "Copy a worksheet."
    },
    "move_worksheet": {
        "required": [
            "sheet_name",
            "position"
        ],
        "optional": [],
        "desc": "Move a worksheet to a different position."
    },
    "hide_worksheet": {
        "required": [
            "sheet_name"
        ],
        "optional": [],
        "desc": "Hide a worksheet."
    },
    "show_worksheet": {
        "required": [
            "sheet_name"
        ],
        "optional": [],
        "desc": "Show a hidden worksheet."
    },
    "write_cell": {
        "required": [
            "sheet_name",
            "cell",
            "value"
        ],
        "optional": [],
        "desc": "Write value to a cell."
    },
    "write_range": {
        "required": [
            "sheet_name",
            "range_addr",
            "values"
        ],
        "optional": [],
        "desc": "Write values to a range."
    },
    "read_cell": {
        "required": [
            "sheet_name",
            "cell"
        ],
        "optional": [],
        "desc": "Read value from a cell."
    },
    "read_range": {
        "required": [
            "sheet_name",
            "range_addr"
        ],
        "optional": [],
        "desc": "Read values from a range."
    },
    "copy_paste_cells": {
        "required": [
            "sheet_name",
            "source_range",
            "dest_range"
        ],
        "optional": [],
        "desc": "Copy and paste cells."
    },
    "clear_contents": {
        "required": [
            "sheet_name",
            "range_addr"
        ],
        "optional": [],
        "desc": "Clear cell contents."
    },
    "find_and_replace": {
        "required": [
            "sheet_name",
            "find_text",
            "replace_text"
        ],
        "optional": [],
        "desc": "Find and replace in worksheet."
    },
    "write_formula": {
        "required": [
            "sheet_name",
            "cell",
            "formula"
        ],
        "optional": [],
        "desc": "Write formula to a cell."
    },
    "use_function": {
        "required": [
            "sheet_name",
            "cell",
            "function_name",
            "range_addr"
        ],
        "optional": [],
        "desc": "Use common function (SUM, AVERAGE, IF, etc.)."
    },
    "use_vlookup": {
        "required": [
            "sheet_name",
            "cell",
            "lookup_value",
            "table_array",
            "col_index"
        ],
        "optional": [
            "exact_match"
        ],
        "desc": "Use VLOOKUP function."
    },
    "set_reference_type": {
        "required": [
            "sheet_name",
            "cell",
            "formula"
        ],
        "optional": [
            "absolute"
        ],
        "desc": "Set formula with absolute/relative references."
    },
    "use_array_formula": {
        "required": [
            "sheet_name",
            "range_addr",
            "formula"
        ],
        "optional": [],
        "desc": "Apply array formula."
    },
    "set_number_format": {
        "required": [
            "sheet_name",
            "range_addr",
            "format_code"
        ],
        "optional": [],
        "desc": "Set number format."
    },
    "set_cell_color": {
        "required": [
            "sheet_name",
            "range_addr",
            "r",
            "g",
            "b"
        ],
        "optional": [],
        "desc": "Set cell background color."
    },
    "set_font_color": {
        "required": [
            "sheet_name",
            "range_addr",
            "r",
            "g",
            "b"
        ],
        "optional": [],
        "desc": "Set font color."
    },
    "set_borders": {
        "required": [
            "sheet_name",
            "range_addr"
        ],
        "optional": [
            "border_style"
        ],
        "desc": "Set cell borders."
    },
    "set_alignment": {
        "required": [
            "sheet_name",
            "range_addr"
        ],
        "optional": [
            "horizontal",
            "vertical"
        ],
        "desc": "Set cell alignment."
    },
    "set_wrap_text": {
        "required": [
            "sheet_name",
            "range_addr"
        ],
        "optional": [
            "wrap"
        ],
        "desc": "Set text wrapping."
    },
    "merge_cells": {
        "required": [
            "sheet_name",
            "range_addr"
        ],
        "optional": [],
        "desc": "Merge cells."
    },
    "set_column_width": {
        "required": [
            "sheet_name",
            "column",
            "width"
        ],
        "optional": [],
        "desc": "Set column width."
    },
    "set_row_height": {
        "required": [
            "sheet_name",
            "row",
            "height"
        ],
        "optional": [],
        "desc": "Set row height."
    },
    "conditional_formatting": {
        "required": [
            "sheet_name",
            "range_addr",
            "condition_type"
        ],
        "optional": [],
        "desc": "Apply conditional formatting."
    },
    "convert_to_table": {
        "required": [
            "sheet_name",
            "range_addr"
        ],
        "optional": [
            "table_name"
        ],
        "desc": "Convert range to table."
    },
    "add_total_row": {
        "required": [
            "sheet_name",
            "table_name"
        ],
        "optional": [],
        "desc": "Add total row to table."
    },
    "apply_table_style": {
        "required": [
            "sheet_name",
            "table_name",
            "style_name"
        ],
        "optional": [],
        "desc": "Apply style to table."
    },
    "filter_table": {
        "required": [
            "sheet_name",
            "table_name",
            "column",
            "criteria"
        ],
        "optional": [],
        "desc": "Filter table."
    },
    "sort_table": {
        "required": [
            "sheet_name",
            "table_name",
            "column"
        ],
        "optional": [
            "ascending"
        ],
        "desc": "Sort table."
    },
    "insert_image": {
        "required": [
            "sheet_name",
            "image_path",
            "cell"
        ],
        "optional": [
            "width",
            "height"
        ],
        "desc": "Insert image in worksheet."
    },
    "resize_image": {
        "required": [
            "sheet_name",
            "image_index",
            "width",
            "height"
        ],
        "optional": [],
        "desc": "Resize image."
    },
    "position_image": {
        "required": [
            "sheet_name",
            "image_index",
            "left",
            "top"
        ],
        "optional": [],
        "desc": "Position image."
    },
    "anchor_image_to_cell": {
        "required": [
            "sheet_name",
            "image_index",
            "cell"
        ],
        "optional": [],
        "desc": "Anchor image to cell."
    },
    "insert_logo_watermark": {
        "required": [
            "sheet_name",
            "image_path"
        ],
        "optional": [],
        "desc": "Insert logo/watermark."
    },
    "create_chart": {
        "required": [
            "sheet_name",
            "chart_type",
            "source_range"
        ],
        "optional": [
            "chart_title"
        ],
        "desc": "Create chart."
    },
    "modify_chart_data": {
        "required": [
            "sheet_name",
            "chart_index",
            "new_range"
        ],
        "optional": [],
        "desc": "Modify chart data source."
    },
    "customize_chart_title": {
        "required": [
            "sheet_name",
            "chart_index",
            "title"
        ],
        "optional": [],
        "desc": "Customize chart title."
    },
    "customize_chart_legend": {
        "required": [
            "sheet_name",
            "chart_index"
        ],
        "optional": [
            "position"
        ],
        "desc": "Customize chart legend."
    },
    "modify_chart_axes": {
        "required": [
            "sheet_name",
            "chart_index"
        ],
        "optional": [
            "x_title",
            "y_title"
        ],
        "desc": "Modify chart axes."
    },
    "change_chart_colors": {
        "required": [
            "sheet_name",
            "chart_index",
            "color_scheme"
        ],
        "optional": [],
        "desc": "Change chart colors and style."
    },
    "move_resize_chart": {
        "required": [
            "sheet_name",
            "chart_index"
        ],
        "optional": [
            "left",
            "top",
            "width",
            "height"
        ],
        "desc": "Move and resize chart."
    },
    "create_pivot_table": {
        "required": [
            "source_sheet",
            "source_range",
            "dest_sheet",
            "dest_cell",
            "table_name"
        ],
        "optional": [],
        "desc": "Create pivot table."
    },
    "set_pivot_fields": {
        "required": [
            "sheet_name",
            "pivot_table_name"
        ],
        "optional": [
            "row_fields",
            "column_fields",
            "value_fields"
        ],
        "desc": "Set pivot table fields."
    },
    "apply_pivot_filter": {
        "required": [
            "sheet_name",
            "pivot_table_name",
            "field",
            "values"
        ],
        "optional": [],
        "desc": "Apply filter to pivot table."
    },
    "change_pivot_calculation": {
        "required": [
            "sheet_name",
            "pivot_table_name",
            "field",
            "function"
        ],
        "optional": [],
        "desc": "Change pivot table calculation."
    },
    "refresh_pivot_table": {
        "required": [
            "sheet_name",
            "pivot_table_name"
        ],
        "optional": [],
        "desc": "Refresh pivot table data."
    },
    "sort_ascending": {
        "required": [
            "sheet_name",
            "range_addr"
        ],
        "optional": [
            "key_column"
        ],
        "desc": "Sort range in ascending order."
    },
    "sort_descending": {
        "required": [
            "sheet_name",
            "range_addr"
        ],
        "optional": [
            "key_column"
        ],
        "desc": "Sort range in descending order."
    },
    "apply_autofilter": {
        "required": [
            "sheet_name",
            "range_addr"
        ],
        "optional": [],
        "desc": "Apply auto filter."
    },
    "create_advanced_filter": {
        "required": [
            "sheet_name",
            "data_range",
            "criteria_range"
        ],
        "optional": [],
        "desc": "Create advanced filter."
    },
    "protect_worksheet": {
        "required": [
            "sheet_name"
        ],
        "optional": [
            "password"
        ],
        "desc": "Protect worksheet."
    },
    "protect_workbook": {
        "required": [],
        "optional": [
            "password"
        ],
        "desc": "Protect workbook structure."
    },
    "set_workbook_password": {
        "required": [
            "password"
        ],
        "optional": [],
        "desc": "Set workbook password."
    },
    "unprotect_worksheet": {
        "required": [
            "sheet_name"
        ],
        "optional": [
            "password"
        ],
        "desc": "Remove worksheet protection."
    },
    "create_named_range": {
        "required": [
            "name",
            "sheet_name",
            "range_addr"
        ],
        "optional": [],
        "desc": "Create named range."
    },
    "use_named_range_in_formula": {
        "required": [
            "sheet_name",
            "cell",
            "range_name"
        ],
        "optional": [
            "function"
        ],
        "desc": "Use named range in formula."
    },
    "delete_named_range": {
        "required": [
            "name"
        ],
        "optional": [],
        "desc": "Delete named range."
    },
    "create_dropdown_list": {
        "required": [
            "sheet_name",
            "range_addr",
            "values"
        ],
        "optional": [],
        "desc": "Create dropdown list."
    },
    "set_validation_rules": {
        "required": [
            "sheet_name",
            "range_addr",
            "validation_type",
            "formula1"
        ],
        "optional": [
            "formula2"
        ],
        "desc": "Set data validation rules."
    },
    "remove_validation": {
        "required": [
            "sheet_name",
            "range_addr"
        ],
        "optional": [],
        "desc": "Remove data validation."
    },
    "configure_print_settings": {
        "required": [
            "sheet_name"
        ],
        "optional": [
            "orientation",
            "paper_size"
        ],
        "desc": "Configure print settings."
    },
    "set_print_area": {
        "required": [
            "sheet_name",
            "range_addr"
        ],
        "optional": [],
        "desc": "Set print area."
    },
    "print_preview": {
        "required": [
            "sheet_name"
        ],
        "optional": [],
        "desc": "Show print preview."
    },
    "group_rows_columns": {
        "required": [
            "sheet_name",
            "range_addr"
        ],
        "optional": [
            "is_rows"
        ],
        "desc": "Group rows or columns."
    },
    "freeze_panes": {
        "required": [
            "sheet_name",
            "cell"
        ],
        "optional": [],
        "desc": "Freeze panes."
    },
    "split_window": {
        "required": [
            "sheet_name"
        ],
        "optional": [
            "horizontal_split",
            "vertical_split"
        ],
        "desc": "Split window."
    },
    "create_sparklines": {
        "required": [
            "sheet_name",
            "data_range",
            "location_range"
        ],
        "optional": [
            "sparkline_type"
        ],
        "desc": "Create sparklines."
    },
    "scenario_analysis": {
        "required": [
            "sheet_name",
            "scenario_name",
            "changing_cells",
            "values"
        ],
        "optional": [],
        "desc": "Create scenario for analysis."
    },
    "goal_seek": {
        "required": [
            "sheet_name",
            "set_cell",
            "to_value",
            "by_changing_cell"
        ],
        "optional": [],
        "desc": "Perform goal seek."
    },
    "use_solver": {
        "required": [
            "sheet_name"
        ],
        "optional": [],
        "desc": "Use Solver add-in."
    },
    "consolidate_data": {
        "required": [
            "dest_sheet",
            "dest_range",
            "source_ranges"
        ],
        "optional": [
            "function"
        ],
        "desc": "Consolidate data from multiple ranges."
    },
    "create_subtotals": {
        "required": [
            "sheet_name",
            "range_addr",
            "group_by"
        ],
        "optional": [
            "function"
        ],
        "desc": "Create automatic subtotals."
    },
    "import_csv": {
        "required": [
            "sheet_name",
            "csv_path"
        ],
        "optional": [
            "dest_cell"
        ],
        "desc": "Import CSV data."
    },
    "insert_hyperlink": {
        "required": [
            "sheet_name",
            "cell",
            "url"
        ],
        "optional": [
            "display_text"
        ],
        "desc": "Insert hyperlink."
    },
    "insert_comment": {
        "required": [
            "sheet_name",
            "cell",
            "comment_text"
        ],
        "optional": [],
        "desc": "Insert comment/note."
    },
    "use_3d_reference": {
        "required": [
            "dest_sheet",
            "dest_cell",
            "first_sheet",
            "last_sheet",
            "cell_ref"
        ],
        "optional": [
            "function"
        ],
        "desc": "Use 3D reference across sheets."
    },
    "export_to_json": {
        "required": [
            "sheet_name",
            "range_addr",
            "output_path"
        ],
        "optional": [],
        "desc": "Export range to JSON."
    }
}

POWERPOINT_TOOLS_CONFIG = {
    "create_presentation": {
        "required": [],
        "optional": [],
        "desc": "Create a new presentation."
    },
    "open_presentation": {
        "required": [
            "file_path"
        ],
        "optional": [],
        "desc": "Open an existing presentation."
    },
    "save_presentation": {
        "required": [],
        "optional": [
            "file_path"
        ],
        "desc": "Save the current presentation."
    },
    "close_presentation": {
        "required": [],
        "optional": [
            "save_changes"
        ],
        "desc": "Close the current presentation."
    },
    "export_to_pdf": {
        "required": [
            "output_path"
        ],
        "optional": [],
        "desc": "Export presentation to PDF."
    },
    "save_as": {
        "required": [
            "file_path"
        ],
        "optional": [
            "file_format"
        ],
        "desc": "Save presentation with different format."
    },
    "create_from_template": {
        "required": [
            "template_path"
        ],
        "optional": [],
        "desc": "Create presentation from template."
    },
    "save_as_template": {
        "required": [
            "template_path"
        ],
        "optional": [],
        "desc": "Save presentation as template."
    },
    "apply_template": {
        "required": [
            "template_path"
        ],
        "optional": [],
        "desc": "Apply template to existing presentation."
    },
    "create_custom_slide_master": {
        "required": [
            "master_name"
        ],
        "optional": [],
        "desc": "Create custom slide master."
    },
    "add_slide": {
        "required": [],
        "optional": [
            "layout"
        ],
        "desc": "Add a new slide."
    },
    "delete_slide": {
        "required": [
            "slide_index"
        ],
        "optional": [],
        "desc": "Delete a slide."
    },
    "duplicate_slide": {
        "required": [
            "slide_index"
        ],
        "optional": [],
        "desc": "Duplicate a slide."
    },
    "move_slide": {
        "required": [
            "slide_index",
            "new_position"
        ],
        "optional": [],
        "desc": "Move slide to new position."
    },
    "apply_slide_layout": {
        "required": [
            "slide_index",
            "layout"
        ],
        "optional": [],
        "desc": "Apply layout to slide."
    },
    "hide_show_slide": {
        "required": [
            "slide_index"
        ],
        "optional": [
            "hidden"
        ],
        "desc": "Hide or show a slide."
    },
    "add_textbox": {
        "required": [
            "slide_index",
            "text",
            "left",
            "top",
            "width",
            "height"
        ],
        "optional": [],
        "desc": "Add text box to slide."
    },
    "modify_title": {
        "required": [
            "slide_index",
            "title_text"
        ],
        "optional": [],
        "desc": "Modify slide title."
    },
    "modify_body_text": {
        "required": [
            "slide_index",
            "body_text"
        ],
        "optional": [],
        "desc": "Modify slide body text."
    },
    "add_bullets": {
        "required": [
            "slide_index",
            "bullet_points"
        ],
        "optional": [],
        "desc": "Add bullet points to slide."
    },
    "add_numbered_list": {
        "required": [
            "slide_index",
            "items"
        ],
        "optional": [],
        "desc": "Add numbered list to slide."
    },
    "format_text": {
        "required": [
            "slide_index",
            "shape_index"
        ],
        "optional": [
            "font_name",
            "font_size",
            "bold",
            "italic",
            "color_rgb"
        ],
        "desc": "Format text in shape."
    },
    "insert_image": {
        "required": [
            "slide_index",
            "image_path",
            "left",
            "top"
        ],
        "optional": [
            "width",
            "height"
        ],
        "desc": "Insert image on slide."
    },
    "resize_image": {
        "required": [
            "slide_index",
            "shape_index",
            "width",
            "height"
        ],
        "optional": [],
        "desc": "Resize image."
    },
    "reposition_image": {
        "required": [
            "slide_index",
            "shape_index",
            "left",
            "top"
        ],
        "optional": [],
        "desc": "Reposition image."
    },
    "insert_video": {
        "required": [
            "slide_index",
            "video_path",
            "left",
            "top",
            "width",
            "height"
        ],
        "optional": [],
        "desc": "Insert video on slide."
    },
    "insert_audio": {
        "required": [
            "slide_index",
            "audio_path",
            "left",
            "top"
        ],
        "optional": [],
        "desc": "Insert audio on slide."
    },
    "insert_shape": {
        "required": [
            "slide_index",
            "shape_type",
            "left",
            "top",
            "width",
            "height"
        ],
        "optional": [],
        "desc": "Insert shape on slide."
    },
    "modify_fill_color": {
        "required": [
            "slide_index",
            "shape_index",
            "r",
            "g",
            "b"
        ],
        "optional": [],
        "desc": "Modify shape fill color."
    },
    "modify_outline": {
        "required": [
            "slide_index",
            "shape_index"
        ],
        "optional": [
            "color_rgb",
            "weight"
        ],
        "desc": "Modify shape outline."
    },
    "group_shapes": {
        "required": [
            "slide_index",
            "shape_indices"
        ],
        "optional": [],
        "desc": "Group multiple shapes."
    },
    "ungroup_shapes": {
        "required": [
            "slide_index",
            "group_index"
        ],
        "optional": [],
        "desc": "Ungroup shapes."
    },
    "insert_table": {
        "required": [
            "slide_index",
            "rows",
            "cols",
            "left",
            "top",
            "width",
            "height"
        ],
        "optional": [],
        "desc": "Insert table on slide."
    },
    "fill_table_cell": {
        "required": [
            "slide_index",
            "table_index",
            "row",
            "col",
            "text"
        ],
        "optional": [],
        "desc": "Fill table cell with text."
    },
    "merge_table_cells": {
        "required": [
            "slide_index",
            "table_index",
            "start_row",
            "start_col",
            "end_row",
            "end_col"
        ],
        "optional": [],
        "desc": "Merge table cells."
    },
    "split_table_cell": {
        "required": [
            "slide_index",
            "table_index",
            "row",
            "col",
            "num_rows",
            "num_cols"
        ],
        "optional": [],
        "desc": "Split table cell."
    },
    "apply_table_style": {
        "required": [
            "slide_index",
            "table_index",
            "style_id"
        ],
        "optional": [],
        "desc": "Apply style to table."
    },
    "format_table_borders": {
        "required": [
            "slide_index",
            "table_index",
            "color_rgb",
            "weight"
        ],
        "optional": [],
        "desc": "Format table borders."
    },
    "insert_chart": {
        "required": [
            "slide_index",
            "chart_type",
            "left",
            "top",
            "width",
            "height"
        ],
        "optional": [],
        "desc": "Insert chart on slide."
    },
    "link_excel_chart": {
        "required": [
            "slide_index",
            "excel_path",
            "left",
            "top",
            "width",
            "height"
        ],
        "optional": [],
        "desc": "Insert chart linked to Excel."
    },
    "modify_chart_data": {
        "required": [
            "slide_index",
            "chart_index",
            "data"
        ],
        "optional": [],
        "desc": "Modify chart data."
    },
    "customize_chart_style": {
        "required": [
            "slide_index",
            "chart_index",
            "style_id"
        ],
        "optional": [],
        "desc": "Customize chart style."
    },
    "add_entrance_animation": {
        "required": [
            "slide_index",
            "shape_index",
            "effect_type"
        ],
        "optional": [],
        "desc": "Add entrance animation."
    },
    "add_exit_animation": {
        "required": [
            "slide_index",
            "shape_index",
            "effect_type"
        ],
        "optional": [],
        "desc": "Add exit animation."
    },
    "set_animation_order": {
        "required": [
            "slide_index",
            "animation_index",
            "new_order"
        ],
        "optional": [],
        "desc": "Set animation order."
    },
    "configure_animation_timing": {
        "required": [
            "slide_index",
            "animation_index"
        ],
        "optional": [
            "duration",
            "delay"
        ],
        "desc": "Configure animation timing."
    },
    "apply_transition": {
        "required": [
            "slide_index",
            "transition_type"
        ],
        "optional": [],
        "desc": "Apply transition to slide."
    },
    "set_transition_duration": {
        "required": [
            "slide_index",
            "duration"
        ],
        "optional": [],
        "desc": "Set transition duration."
    },
    "apply_transition_to_all": {
        "required": [
            "transition_type"
        ],
        "optional": [
            "duration"
        ],
        "desc": "Apply transition to all slides."
    },
    "apply_theme": {
        "required": [
            "theme_path"
        ],
        "optional": [],
        "desc": "Apply theme to presentation."
    },
    "modify_color_scheme": {
        "required": [
            "color_scheme_index"
        ],
        "optional": [],
        "desc": "Modify color scheme."
    },
    "modify_theme_fonts": {
        "required": [
            "heading_font",
            "body_font"
        ],
        "optional": [],
        "desc": "Modify theme fonts."
    },
    "set_background": {
        "required": [
            "slide_index"
        ],
        "optional": [
            "color_rgb",
            "image_path"
        ],
        "desc": "Set slide background."
    },
    "apply_slide_master": {
        "required": [
            "master_index"
        ],
        "optional": [],
        "desc": "Apply slide master."
    },
    "add_speaker_notes": {
        "required": [
            "slide_index",
            "notes_text"
        ],
        "optional": [],
        "desc": "Add speaker notes to slide."
    },
    "read_speaker_notes": {
        "required": [
            "slide_index"
        ],
        "optional": [],
        "desc": "Read speaker notes from slide."
    },
    "add_comment": {
        "required": [
            "slide_index",
            "text",
            "left",
            "top"
        ],
        "optional": [
            "author"
        ],
        "desc": "Add comment to slide."
    },
    "start_presenter_mode": {
        "required": [],
        "optional": [],
        "desc": "Start presenter mode."
    },
    "set_slide_timing": {
        "required": [
            "slide_index",
            "advance_time"
        ],
        "optional": [],
        "desc": "Set automatic slide timing."
    },
    "record_slideshow": {
        "required": [
            "output_path"
        ],
        "optional": [],
        "desc": "Record slideshow with narration."
    },
    "insert_smartart": {
        "required": [
            "slide_index",
            "layout",
            "left",
            "top",
            "width",
            "height"
        ],
        "optional": [],
        "desc": "Insert SmartArt."
    },
    "insert_ole_object": {
        "required": [
            "slide_index",
            "file_path",
            "left",
            "top",
            "width",
            "height"
        ],
        "optional": [],
        "desc": "Insert OLE object (Excel, equations, etc.)."
    },
    "create_section_zoom": {
        "required": [
            "slide_index",
            "section_name",
            "left",
            "top"
        ],
        "optional": [],
        "desc": "Create section zoom."
    },
    "insert_hyperlink": {
        "required": [
            "slide_index",
            "shape_index",
            "url"
        ],
        "optional": [
            "target_slide"
        ],
        "desc": "Insert hyperlink to shape."
    },
    "add_action_trigger": {
        "required": [
            "slide_index",
            "shape_index",
            "action_type"
        ],
        "optional": [],
        "desc": "Add action and trigger to shape."
    },
    "export_to_video": {
        "required": [
            "output_path"
        ],
        "optional": [
            "frame_rate"
        ],
        "desc": "Export presentation to video."
    },
    "add_captions": {
        "required": [
            "slide_index",
            "caption_text"
        ],
        "optional": [],
        "desc": "Add captions for accessibility."
    },
    "compare_presentations": {
        "required": [
            "other_path"
        ],
        "optional": [],
        "desc": "Compare two presentations."
    }
}

OUTLOOK_TOOLS_CONFIG = {
    "send_email": {
        "required": ["to", "subject", "body"],
        "optional": ["cc", "bcc", "importance"],
        "desc": "Envoie un email via Outlook"
    },
    "send_with_attachments": {
        "required": ["to", "subject", "body", "attachments"],
        "optional": ["cc", "bcc", "importance"],
        "desc": "Envoie un email avec pièces jointes"
    },
    "read_email": {
        "required": ["email_entry_id"],
        "optional": [],
        "desc": "Lit les détails d'un email"
    },
    "reply_to_email": {
        "required": ["email_entry_id", "body"],
        "optional": ["send_immediately"],
        "desc": "Répond à un email"
    },
    "reply_all_to_email": {
        "required": ["email_entry_id", "body"],
        "optional": ["send_immediately"],
        "desc": "Répond à tous les destinataires"
    },
    "forward_email": {
        "required": ["email_entry_id", "to"],
        "optional": ["body", "send_immediately"],
        "desc": "Transfère un email"
    },
    "mark_as_read": {
        "required": ["email_entry_id"],
        "optional": [],
        "desc": "Marque un email comme lu"
    },
    "mark_as_unread": {
        "required": ["email_entry_id"],
        "optional": [],
        "desc": "Marque un email comme non lu"
    },
    "flag_email": {
        "required": ["email_entry_id"],
        "optional": ["flag_status"],
        "desc": "Ajoute un drapeau sur un email"
    },
    "delete_email": {
        "required": ["email_entry_id"],
        "optional": [],
        "desc": "Supprime un email"
    },
    "move_email_to_folder": {
        "required": ["email_entry_id", "folder_path"],
        "optional": [],
        "desc": "Déplace un email vers un dossier"
    },
    "search_emails": {
        "required": [],
        "optional": ["folder_name", "subject", "sender", "body_contains", 
                    "start_date", "end_date", "unread_only", "max_results"],
        "desc": "Recherche des emails"
    },
    "add_attachment": {
        "required": ["email_entry_id", "file_path"],
        "optional": ["display_name"],
        "desc": "Ajoute une pièce jointe"
    },
    "list_attachments": {
        "required": ["email_entry_id"],
        "optional": [],
        "desc": "Liste les pièces jointes"
    },
    "save_attachment": {
        "required": ["email_entry_id", "attachment_index", "save_path"],
        "optional": [],
        "desc": "Sauvegarde une pièce jointe"
    },
    "remove_attachment": {
        "required": ["email_entry_id", "attachment_index"],
        "optional": [],
        "desc": "Supprime une pièce jointe"
    },
    "create_new_message": {
        "required": [],
        "optional": [],
        "desc": "Crée un nouveau brouillon"
    },
    "create_folder": {
        "required": ["folder_name"],
        "optional": ["parent_folder"],
        "desc": "Crée un dossier"
    },
    "delete_folder": {
        "required": ["folder_path"],
        "optional": [],
        "desc": "Supprime un dossier"
    },
    "rename_folder": {
        "required": ["folder_path", "new_name"],
        "optional": [],
        "desc": "Renomme un dossier"
    },
    "move_folder": {
        "required": ["folder_path", "destination_path"],
        "optional": [],
        "desc": "Déplace un dossier"
    },
    "list_folders": {
        "required": [],
        "optional": ["parent_folder", "recursive"],
        "desc": "Liste les dossiers"
    },
    "get_folder_item_count": {
        "required": ["folder_path"],
        "optional": [],
        "desc": "Compte les éléments d'un dossier"
    },
    "get_unread_count": {
        "required": [],
        "optional": ["folder_path"],
        "desc": "Compte les messages non lus"
    },
    "create_appointment": {
        "required": ["subject", "start_time", "end_time"],
        "optional": ["location", "body", "reminder_minutes", "busy_status"],
        "desc": "Crée un rendez-vous"
    },
    "create_recurring_event": {
        "required": ["subject", "start_time", "end_time", "recurrence_type"],
        "optional": ["interval", "occurrences", "end_date", "location", "body"],
        "desc": "Crée un événement récurrent"
    },
    "read_appointment": {
        "required": ["appointment_entry_id"],
        "optional": [],
        "desc": "Lit un rendez-vous"
    },
    "modify_appointment": {
        "required": ["appointment_entry_id"],
        "optional": ["subject", "start_time", "end_time", "location", "body"],
        "desc": "Modifie un rendez-vous"
    },
    "delete_appointment": {
        "required": ["appointment_entry_id"],
        "optional": [],
        "desc": "Supprime un rendez-vous"
    },
    "search_appointments": {
        "required": [],
        "optional": ["subject", "location", "start_date", "end_date", "max_results"],
        "desc": "Recherche des rendez-vous"
    },
    "get_appointments_by_date": {
        "required": ["start_date", "end_date"],
        "optional": [],
        "desc": "Obtient les rendez-vous par date"
    },
    "set_reminder": {
        "required": ["appointment_entry_id", "reminder_minutes"],
        "optional": [],
        "desc": "Définit un rappel"
    },
    "set_busy_status": {
        "required": ["appointment_entry_id", "busy_status"],
        "optional": [],
        "desc": "Définit le statut occupé"
    },
    "export_appointment_ics": {
        "required": ["appointment_entry_id", "output_path"],
        "optional": [],
        "desc": "Exporte en ICS"
    },
    "get_calendar_count": {
        "required": [],
        "optional": [],
        "desc": "Compte les rendez-vous"
    },
    "export_to_pdf": {
        "required": ["output_path"],
        "optional": [],
        "desc": "Exporte le calendrier en PDF"
    },
    "create_meeting_request": {
        "required": ["subject", "start_time", "end_time", "required_attendees"],
        "optional": ["optional_attendees", "location", "body"],
        "desc": "Crée une demande de réunion"
    },
    "invite_participants": {
        "required": ["meeting_entry_id", "attendees"],
        "optional": ["required"],
        "desc": "Invite des participants"
    },
    "accept_meeting": {
        "required": ["meeting_entry_id"],
        "optional": [],
        "desc": "Accepte une réunion"
    },
    "decline_meeting": {
        "required": ["meeting_entry_id"],
        "optional": [],
        "desc": "Refuse une réunion"
    },
    "propose_new_time": {
        "required": ["meeting_entry_id", "new_start", "new_end"],
        "optional": [],
        "desc": "Propose un nouveau créneau"
    },
    "cancel_meeting": {
        "required": ["meeting_entry_id"],
        "optional": [],
        "desc": "Annule une réunion"
    },
    "update_meeting": {
        "required": ["meeting_entry_id"],
        "optional": ["subject", "start_time", "end_time", "location"],
        "desc": "Met à jour une réunion"
    },
    "check_availability": {
        "required": ["attendees", "start_time", "end_time"],
        "optional": ["duration_minutes"],
        "desc": "Vérifie la disponibilité"
    },
    "create_contact": {
        "required": ["first_name", "last_name"],
        "optional": ["email", "phone", "company", "job_title"],
        "desc": "Crée un contact"
    },
    "modify_contact": {
        "required": ["contact_entry_id"],
        "optional": ["first_name", "last_name", "email", "phone"],
        "desc": "Modifie un contact"
    },
    "delete_contact": {
        "required": ["contact_entry_id"],
        "optional": [],
        "desc": "Supprime un contact"
    },
    "search_contact": {
        "required": ["search_term"],
        "optional": [],
        "desc": "Recherche un contact"
    },
    "list_all_contacts": {
        "required": [],
        "optional": [],
        "desc": "Liste tous les contacts"
    },
    "create_contact_group": {
        "required": ["group_name"],
        "optional": [],
        "desc": "Crée un groupe de contacts"
    },
    "add_to_contact_group": {
        "required": ["group_entry_id", "contact_email"],
        "optional": [],
        "desc": "Ajoute à un groupe"
    },
    "export_contacts_vcf": {
        "required": ["output_path"],
        "optional": [],
        "desc": "Exporte les contacts en VCF"
    },
    "import_contacts": {
        "required": ["file_path"],
        "optional": [],
        "desc": "Importe des contacts"
    },
    "create_task": {
        "required": ["subject"],
        "optional": ["body", "due_date", "priority"],
        "desc": "Crée une tâche"
    },
    "modify_task": {
        "required": ["task_entry_id"],
        "optional": ["subject", "body", "due_date"],
        "desc": "Modifie une tâche"
    },
    "delete_task": {
        "required": ["task_entry_id"],
        "optional": [],
        "desc": "Supprime une tâche"
    },
    "mark_task_complete": {
        "required": ["task_entry_id"],
        "optional": [],
        "desc": "Marque une tâche terminée"
    },
    "set_task_priority": {
        "required": ["task_entry_id", "priority"],
        "optional": [],
        "desc": "Définit la priorité"
    },
    "set_task_due_date": {
        "required": ["task_entry_id", "due_date"],
        "optional": [],
        "desc": "Définit l'échéance"
    },
    "list_tasks": {
        "required": [],
        "optional": ["completed"],
        "desc": "Liste les tâches"
    },
    "list_accounts": {
        "required": [],
        "optional": [],
        "desc": "Liste les comptes"
    },
    "get_default_account": {
        "required": [],
        "optional": [],
        "desc": "Obtient le compte par défaut"
    },
    "get_inbox_count": {
        "required": [],
        "optional": [],
        "desc": "Compte les messages inbox"
    },
    "create_category": {
        "required": ["name"],
        "optional": ["color"],
        "desc": "Crée une catégorie"
    },
    "list_categories": {
        "required": [],
        "optional": [],
        "desc": "Liste les catégories"
    },
    "apply_category": {
        "required": ["item_entry_id", "category"],
        "optional": [],
        "desc": "Applique une catégorie"
    },
    "com_operation": {
        "required": ["operation_name"],
        "optional": [],
        "desc": "Opération COM personnalisée"
    }
}
