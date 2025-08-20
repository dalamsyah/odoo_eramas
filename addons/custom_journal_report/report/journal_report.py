from odoo import models

class JournalEntryReport(models.AbstractModel):
    _name = "report.custom_journal_report.journal_report_template"
    _description = "Custom Journal Entries Report"

    def _get_report_values(self, docids, data=None):
        docs = self.env['account.move'].browse(docids)
        return {
            'doc_ids': docids,
            'doc_model': 'account.move',
            'docs': docs,
        }
