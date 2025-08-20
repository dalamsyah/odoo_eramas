from odoo import models

class TrialBalanceReport(models.AbstractModel):
    _inherit = "account.move.line"
    
    _name = 'report.custom_trial_balance_report.trial_balance_template'
    _description = 'Trial Balance PDF'

    def _get_report_base_filename(self):
        return "Trial Balance - %s" % self.display_name

    def _get_report_values(self, docids, data=None):
        lines = self.env['account.move.line'].read_group(
            [('parent_state', '=', 'posted')],
            ['account_id', 'debit', 'credit'],
            ['account_id']
        )
        result = []
        for line in lines:
            account = self.env['account.account'].browse(line['account_id'][0])
            balance = line['debit'] - line['credit']
            result.append({
                'code': account.code,
                'name': account.name,
                'debit': line['debit'],
                'credit': line['credit'],
                'balance': balance,
            })
        return {
            'doc_ids': docids,
            'doc_model': 'account.move.line',
            'docs': result,
        }
