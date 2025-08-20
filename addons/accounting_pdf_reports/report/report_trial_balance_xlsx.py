# -*- coding: utf-8 -*-
from odoo import models, _
from odoo.exceptions import UserError
import io
import base64
import time
import xlsxwriter

class ReportTrialBalanceXlsx(models.AbstractModel):
    _name = 'report.accounting_pdf_reports.report_trial_balance_xlsx'
    _inherit = 'report.report_xlsx.abstract'
    _description = 'Trial Balance Report Xlsx'

    def _get_accounts(self, accounts, display_account):
        """ compute the balance, debit and credit for the provided accounts
            :Arguments:
                `accounts`: list of accounts record,
                `display_account`: it's used to display either all accounts or those accounts which balance is > 0
            :Returns a list of dictionary of Accounts with following key and value
                `name`: Account name,
                `code`: Account code,
                `credit`: total amount of credit,
                `debit`: total amount of debit,
                `balance`: total amount of balance,
        """

        account_result = {}
        # Prepare sql query base on selected parameters from wizard
        tables, where_clause, where_params = self.env['account.move.line']._query_get()
        tables = tables.replace('"','')
        if not tables:
            tables = 'account_move_line'
        wheres = [""]
        if where_clause.strip():
            wheres.append(where_clause.strip())
        filters = " AND ".join(wheres)
        # compute the balance, debit and credit for the provided accounts
        request = ("SELECT account_id AS id, SUM(debit) AS debit, SUM(credit) AS credit, "
                   "(SUM(debit) - SUM(credit)) AS balance" +\
                   " FROM " + tables + " WHERE account_id IN %s " + filters + " GROUP BY account_id")
        params = (tuple(accounts.ids),) + tuple(where_params)
        self.env.cr.execute(request, params)
        for row in self.env.cr.dictfetchall():
            account_result[row.pop('id')] = row

        account_res = []
        for account in accounts:
            res = dict((fn, 0.0) for fn in ['credit', 'debit', 'balance'])
            currency = account.currency_id and account.currency_id or account.company_id.currency_id
            res['code'] = account.code
            res['name'] = account.name
            if account.id in account_result:
                res['debit'] = account_result[account.id].get('debit')
                res['credit'] = account_result[account.id].get('credit')
                res['balance'] = account_result[account.id].get('balance')
            if display_account == 'all':
                account_res.append(res)
            if display_account == 'not_zero' and not currency.is_zero(res['balance']):
                account_res.append(res)
            if display_account == 'movement' and (not currency.is_zero(res['debit']) or not currency.is_zero(res['credit'])):
                account_res.append(res)
        return account_res

    def _get_beginning_balances(self, accounts, date_from):
        account_result = {}
        if not date_from:
            return {}

        # Query untuk ambil saldo sampai sebelum date_from
        query = """
            SELECT account_id AS id, 
                SUM(debit) AS debit, 
                SUM(credit) AS credit, 
                (SUM(debit) - SUM(credit)) AS balance
            FROM account_move_line
            WHERE account_id IN %s
            AND date < %s
            GROUP BY account_id
        """
        params = (tuple(accounts.ids), date_from)
        self.env.cr.execute(query, params)
        for row in self.env.cr.dictfetchall():
            account_result[row.pop('id')] = row
        return account_result

    def generate_xlsx_report(self, workbook, data, wizard_records):
        sheet = workbook.add_worksheet('Trial Balance')
        print("ini generate_xlsx_report")
        print(data)

        if not data.get('form') or not self.env.context.get('active_model'):
            raise UserError(_("Form content is missing, this report cannot be printed."))

        model = self.env.context.get('active_model')
        docs = self.env[model].browse(self.env.context.get('active_ids', []))
        display_account = data['form'].get('display_account')
        accounts = docs if model == 'account.account' else self.env['account.account'].search([])
        context = data['form'].get('used_context') or {}

        analytic_accounts = []
        if data['form'].get('analytic_account_ids'):
            analytic_account_ids = self.env['account.analytic.account'].browse(data['form'].get('analytic_account_ids'))
            context['analytic_account_ids'] = analytic_account_ids.ids
            analytic_accounts = [account.name for account in analytic_account_ids]

        account_res = self.with_context(context)._get_accounts(accounts, display_account)
        codes = []
        if data['form'].get('journal_ids', False):
            codes = [journal.code for journal in
                    self.env['account.journal'].search(
                        [('id', 'in', data['form']['journal_ids'])])]

        # === Mulai bikin file Excel ===
        output = io.BytesIO()
        
        # Styling
        bold = workbook.add_format({'bold': True})
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14})
        sub_title_fmt = workbook.add_format({'italic': True, 'font_size': 10})
        money_fmt = workbook.add_format({'num_format': '#,##0.00'})

        # Ambil nama company dari wizard (company_id)
        company = self.env.company.name
        date_from = data['form'].get('date_from')
        date_to = data['form'].get('date_to')

        beginning_balances = self._get_beginning_balances(accounts, date_from)

        # Judul laporan
        sheet.merge_range(0, 0, 0, 4, company, title_fmt)
        sheet.merge_range(1, 0, 1, 4, "Trial Balance", bold)

        # Periode laporan
        if date_from and date_to:
            periode_text = "Periode: %s s/d %s" % (date_from, date_to)
        elif date_from:
            periode_text = "Periode mulai %s" % date_from
        elif date_to:
            periode_text = "Periode sampai %s" % date_to
        else:
            periode_text = "Periode: Semua"

        sheet.merge_range(2, 0, 2, 4, periode_text, sub_title_fmt)

        # Header kolom
        sheet.write(4, 0, 'Code', bold)
        sheet.write(4, 1, 'Account', bold)
        sheet.write(4, 2, 'Beginning Balance', bold)
        sheet.write(4, 3, 'Debit', bold)
        sheet.write(4, 4, 'Credit', bold)
        sheet.write(4, 5, 'Ending Balance', bold)

        row = 5
        for acc in account_res:
            beg_bal = beginning_balances.get(acc['code']) or {}
            beg_balance = beg_bal.get('balance', 0.0)

            ending_balance = beg_balance + acc['debit'] - acc['credit']

            sheet.write(row, 0, acc['code'])
            sheet.write(row, 1, acc['name'])
            sheet.write_number(row, 2, beg_balance, money_fmt)
            sheet.write_number(row, 3, acc['debit'], money_fmt)
            sheet.write_number(row, 4, acc['credit'], money_fmt)
            sheet.write_number(row, 5, ending_balance, money_fmt)
            row += 1

        workbook.close()

        # Simpan sebagai attachment di Odoo
        file_data = base64.b64encode(output.getvalue())
        output.close()

        attachment = self.env['ir.attachment'].create({
            'name': 'trial_balance.xlsx',
            'type': 'binary',
            'datas': file_data,
            'res_model': model,
            'res_id': docs.id if docs else 0,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })

        return {
            'file': attachment.id,
            'file_name': 'trial_balance.xlsx',
        }

    # def generate_xlsx_report(self, workbook, data, wizard_records):
    #     sheet = workbook.add_worksheet('Trial Balance')
    #     print("ini generate_xlsx_report")
    #     print(data)

    #     if not data.get('form') or not self.env.context.get('active_model'):
    #         raise UserError(_("Form content is missing, this report cannot be printed."))

    #     model = self.env.context.get('active_model')
    #     docs = self.env[model].browse(self.env.context.get('active_ids', []))
    #     display_account = data['form'].get('display_account')
    #     accounts = docs if model == 'account.account' else self.env['account.account'].search([])
    #     context = data['form'].get('used_context') or {}

    #     analytic_accounts = []
    #     if data['form'].get('analytic_account_ids'):
    #         analytic_account_ids = self.env['account.analytic.account'].browse(data['form'].get('analytic_account_ids'))
    #         context['analytic_account_ids'] = analytic_account_ids.ids
    #         analytic_accounts = [account.name for account in analytic_account_ids]

    #     account_res = self.with_context(context)._get_accounts(accounts, display_account)
    #     codes = []
    #     if data['form'].get('journal_ids', False):
    #         codes = [journal.code for journal in
    #                  self.env['account.journal'].search(
    #                      [('id', 'in', data['form']['journal_ids'])])]

    #     # === Mulai bikin file Excel ===
    #     output = io.BytesIO()
        
    #     # Styling
    #     bold = workbook.add_format({'bold': True})
    #     money_fmt = workbook.add_format({'num_format': '#,##0.00'})

    #     # Header
    #     sheet.write(0, 0, 'Code', bold)
    #     sheet.write(0, 1, 'Account', bold)
    #     sheet.write(0, 2, 'Debit', bold)
    #     sheet.write(0, 3, 'Credit', bold)
    #     sheet.write(0, 4, 'Balance', bold)

    #     row = 1
    #     for acc in account_res:
    #         sheet.write(row, 0, acc['code'])
    #         sheet.write(row, 1, acc['name'])
    #         sheet.write_number(row, 2, acc['debit'], money_fmt)
    #         sheet.write_number(row, 3, acc['credit'], money_fmt)
    #         sheet.write_number(row, 4, acc['balance'], money_fmt)
    #         row += 1

    #     workbook.close()

    #     # Simpan sebagai attachment di Odoo
    #     file_data = base64.b64encode(output.getvalue())
    #     output.close()

    #     attachment = self.env['ir.attachment'].create({
    #         'name': 'trial_balance.xlsx',
    #         'type': 'binary',
    #         'datas': file_data,
    #         'res_model': model,
    #         'res_id': docs.id if docs else 0,
    #         'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    #     })

    #     return {
    #         'file': attachment.id,
    #         'file_name': 'trial_balance.xlsx',
    #     }
