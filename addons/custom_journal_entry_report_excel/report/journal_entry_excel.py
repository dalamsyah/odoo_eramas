# -*- coding: utf-8 -*-
from odoo import models

class JournalEntryExcel(models.AbstractModel):
    _name = "report.journal_entry_excel.report_journal_entry_excel"
    _inherit = "report.report_xlsx.abstract"
    _description = "Journal Entry Excel Report"

    def generate_xlsx_report(self, workbook, data, moves):
        sheet = workbook.add_worksheet("Journal Entry")
        bold = workbook.add_format({"bold": True})
        money_fmt = workbook.add_format({"num_format": "#,##0.00"})
        title_fmt = workbook.add_format({"bold": True, "font_size": 14})

        row = 0
        col = 0

        for move in moves:
            # --- Header Atas ---
            company_name = move.company_id.name or ""
            sheet.merge_range(row, col, row, col + 6, company_name, title_fmt)
            row += 2

            sheet.write(row, 0, "Journal Number:", bold)
            sheet.write(row, 1, move.name or "")
            row += 1

            sheet.write(row, 0, "Journal Date:", bold)
            sheet.write(row, 1, str(move.date))
            row += 2

            # --- Table Header ---
            sheet.write(row, 0, "Journal", bold)
            sheet.write(row, 1, "Reference", bold)
            sheet.write(row, 2, "Account", bold)
            sheet.write(row, 3, "Label", bold)
            sheet.write(row, 4, "Debit", bold)
            sheet.write(row, 5, "Credit", bold)
            row += 1

            # --- Isi Data ---
            for line in move.line_ids:
                sheet.write(row, 0, move.journal_id.name)
                sheet.write(row, 1, move.ref or "")
                sheet.write(row, 2, f"{line.account_id.code} - {line.account_id.name}")
                sheet.write(row, 3, line.name or "")
                sheet.write_number(row, 4, line.debit, money_fmt)
                sheet.write_number(row, 5, line.credit, money_fmt)
                row += 1

            # kasih jarak antar jurnal kalau banyak
            row += 2
