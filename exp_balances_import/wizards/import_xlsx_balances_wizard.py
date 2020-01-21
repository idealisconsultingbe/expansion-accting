from odoo import api, fields, models, _
import xlrd
import tempfile
import binascii
from _datetime import datetime


def get_date_format(str):
    data = str.split('.')
    str_date = data[1] + "-" + data[0] + "-" + data[2]
    return datetime.strptime(str_date, '%m-%d-%Y').date()


class BalancesXlsxDataWizard(models.TransientModel):
    _name = 'import.balances.data.wizard'
    _description = 'Import Balances xlsx data Wizard'

    xlsx_file = fields.Binary(string="Your xlsx file")

    def import_balances_xlsx_data(self):
        fp = tempfile.NamedTemporaryFile(suffix=".xlsx")
        fp.write(binascii.a2b_base64(self.xlsx_file))
        fp.seek(0)
        workbook = xlrd.open_workbook(fp.name)
        sheet = workbook.sheet_by_index(0)

        col_names = sheet.row_values(0)

        is_supplier = "Code fournisseur" in col_names

        self._cr.autocommit(False)
        try:
            for i in range(1, sheet.nrows):
                row = sheet.row_values(i)

                name_client = row[col_names.index("Nom du client")] if "Nom du Client" in col_names else row[col_names.index("Nom du fournisseur")]
                amount_due = row[col_names.index("Solde dû")]
                new_partner = name_client != ""
                if name_client != "":
                    name_to_save = name_client
                    amount_total = amount_due

                if not new_partner:
                    partner = self.env['res.partner'].search([('name', '=', name_to_save)])
                    document_name = row[col_names.index("Numéro de document")]
                    registration_date = row[col_names.index("Date enregistrement")]
                    date_due = row[col_names.index("Date d'échéance")]
                    ref_number = row[col_names.index("N° de référence partenaire")]

                    account_move = self.env['account.move'].search([('partner_id', '=', partner.id)])
                    if not account_move:
                        account_move = self.env['account.move'].create({
                            'partner_id': partner.id,
                            'date': get_date_format(registration_date),
                            'ref': ref_number,
                            'invoice_date_due': get_date_format(date_due),
                            'amount_total': amount_total,
                        })

                    credit = amount_due if amount_due > 0 else 0
                    if amount_due < 0:
                        debit = abs(amount_due)
                    else:
                        debit = 0

                    final_credit = debit if is_supplier else credit
                    final_debit = credit if is_supplier else debit

                    self.env['account.move.line'].create([{
                        'account_id': partner.property_account_receivable_id.id,
                        'name': int(document_name),
                        'move_id': account_move.id,
                        'partner_id': partner.id,
                        'credit': final_credit,
                        'debit': final_debit,
                    }, {
                        'account_id': partner.property_account_payable_id.id,
                        'move_id': account_move.id,
                        'credit': final_debit if final_debit > 0 else 0,
                        'debit': final_credit if final_credit > 0 else 0,
                    }])

                    # account_move.post()

            self._cr.commit()

        except Exception as e:
            print(e)
            self._cr.rollback()
            raise Warning("Something goes wrong during the import!")
