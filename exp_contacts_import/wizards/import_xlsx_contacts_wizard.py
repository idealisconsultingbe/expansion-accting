from odoo import api, fields, models, _
import xlrd
import tempfile
import binascii


class ContactsXlsxDataWizard(models.TransientModel):
    _name = 'import.contacts.data.wizard'
    _description = 'Import Contacts xlsx data Wizard'

    xlsx_file = fields.Binary(string="Your xlsx file")

    def import_contacts_xlsx_data(self):
        fp = tempfile.NamedTemporaryFile(suffix=".xlsx")
        fp.write(binascii.a2b_base64(self.xlsx_file))
        fp.seek(0)
        workbook = xlrd.open_workbook(fp.name)
        sheet = workbook.sheet_by_index(0)

        # Avoid persist data before the import is finished
        self._cr.autocommit(False)

        col_names = sheet.row_values(0)

        try:
            for i in range(1, sheet.nrows):
                row = sheet.row_values(i)

                # todo generate a uniq external ID
                name = row[col_names.index("Partenaire")]
                partner_code = row[col_names.index("Code Partenaire")]
                vat = row[col_names.index("Numéro d'identification d'entreprise")]
                country_name = row[col_names.index("Pays (destinataire facture)")]
                contact_name = row[col_names.index("Contact")]
                partner_type = row[col_names.index("Type de partenaire")]
                customer_rank = 1 if partner_type == 'Client' else 0
                supplier_rank = 1 if partner_type == 'Fournisseur' else 0
                created = row[col_names.index("Date de création")]
                zip = row[col_names.index("CP (destin. facture)")]
                city = row[col_names.index("Ville destinataire de facture")]
                phone = row[col_names.index("Téléphone 1")]
                street = row[col_names.index("Rue (destinataire facture)")]
                email = row[col_names.index("E-mail")]

                partner = self.env['res.partner'].search([('vat', '=', vat), ('vat', '!=', "")])
                if len(partner) > 1:
                    partner = partner[0]

                if not partner:
                    partner = self.env['res.partner'].create({
                        'company_type': 'company',
                        'name': name,
                        'vat': vat,
                        'customer_rank': customer_rank,
                        'supplier_rank': supplier_rank,
                        'ref': partner_code if customer_rank == 1 else "",
                        'supplier_ref': partner_code if supplier_rank == 1 else "",
                        'street': street if street != "x" else "",
                        'city': city if city != "x" else "",
                        'country_id': self.env['res.country'].search([('name', '=', country_name)]).id,
                        'zip': zip if zip != "x" else "",
                        'phone': phone,
                        'email': email,
                        'create_date': created,
                    })

                    contact = self.env['res.partner'].search([('name', '=', contact_name)])
                    if not contact:
                        contact = self.env['res.partner'].create({'name': contact_name})
                        contact.parent_id = partner.id
                else:
                    partner.write({
                        'ref': partner_code if customer_rank == 1 else partner.ref,
                        'supplier_ref': partner_code if supplier_rank == 1 else partner.supplier_ref
                    })

            self._cr.commit()

        except Exception as e:
            print(e)
            self._cr.rollback()
            raise Warning("Something goes wrong during the import!")
