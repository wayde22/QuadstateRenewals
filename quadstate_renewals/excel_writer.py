import logging

import pandas as pd

from .constants import (
    COMPLETED_BY_DROPDOWN,
    CONTACTED_VIA_DROPDOWN,
    NOTES_DROPDOWN,
    STATE_DROPDOWN,
    STATE_FORMAT,
)


def export_to_excel(df, output_file_path):
    logging.debug(f'Exporting DataFrame to Excel file: {output_file_path}')
    writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')
    workbook = writer.book
    workbook.use_zip64()
    workbook.nan_inf_to_errors = True
    df.to_excel(writer, index=False)
    worksheet = writer.sheets['Sheet1']

    logging.debug('Setting up data validation for dropdowns.')
    state_col_letter = chr(ord('A') + df.columns.get_loc('State'))
    state_range = f'{state_col_letter}2:{state_col_letter}{len(df) + 1}'
    worksheet.data_validation(state_range, {'validate': 'list', 'source': STATE_DROPDOWN})

    contacted_via_col_letter = chr(ord('A') + df.columns.get_loc('Contacted VIA'))
    contacted_via_range = f'{contacted_via_col_letter}2:{contacted_via_col_letter}{len(df) + 1}'
    worksheet.data_validation(contacted_via_range, {'validate': 'list', 'source': CONTACTED_VIA_DROPDOWN})

    notes_col_letter = chr(ord('A') + df.columns.get_loc('Notes Filed'))
    notes_range = f'{notes_col_letter}2:{notes_col_letter}{len(df) + 1}'
    worksheet.data_validation(notes_range, {'validate': 'list', 'source': NOTES_DROPDOWN})

    completed_by_col_letter = chr(ord('A') + df.columns.get_loc('Completed By'))
    completed_by_range = f'{completed_by_col_letter}2:{completed_by_col_letter}{len(df) + 1}'
    worksheet.data_validation(completed_by_range, {'validate': 'list', 'source': COMPLETED_BY_DROPDOWN})

    logging.debug('Setting column widths and formats.')
    for column in df.columns:
        column_width = 16 if column in ['State', 'Contacted VIA', 'Notes Filed', 'Completed By'] else (
            50 if column == 'Notes' else max(df[column].astype(str).map(len).max(), len(column)))
        col_idx = df.columns.get_loc(column)
        worksheet.set_column(col_idx, col_idx, column_width)

    header_format = workbook.add_format({'bold': True, 'bg_color': '#368be9', 'align': 'center', 'valign': 'vcenter'})
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    logging.debug('Applying conditional formatting based on state.')
    data_range = f'A2:{chr(ord("A") + len(df.columns) - 1)}{len(df) + 1}'

    for state, format_spec in STATE_FORMAT.items():
        format_ = workbook.add_format(format_spec)
        worksheet.conditional_format(data_range,
                                     {'type': 'formula',
                                      'criteria': f'${state_col_letter}2="{state}"',
                                      'format': format_})

    center_format = workbook.add_format({'align': 'center'})
    notes_col_index = df.columns.get_loc('Notes Filed')
    contacted_via_col_index = df.columns.get_loc('Contacted VIA')

    for row in range(1, len(df) + 1):
        cell_value = df.iloc[row - 1, notes_col_index]
        worksheet.write(row, notes_col_index, cell_value, center_format)

        cell_value = df.iloc[row - 1, contacted_via_col_index]
        worksheet.write(row, contacted_via_col_index, cell_value, center_format)

    grey_format = workbook.add_format({'bg_color': '#f0f0f0'})
    for row in range(1, len(df) + 1, 2):
        for col in range(len(df.columns)):
            if col == notes_col_index or col == contacted_via_col_index:
                cell_format = workbook.add_format({'bg_color': '#f0f0f0', 'align': 'center'})
            else:
                cell_format = grey_format
            worksheet.write(row, col, df.iloc[row - 1, col], cell_format)

    writer.close()
    logging.info(f'DataFrame exported to {output_file_path}')
    return True

