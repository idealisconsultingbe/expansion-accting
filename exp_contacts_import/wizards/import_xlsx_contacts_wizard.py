from odoo import api, fields, models, _
import xlrd
import tempfile
import binascii

_ref_vat = {
    'at': 'ATU12345675',
    'be': 'BE0477472701',
    'bg': 'BG1234567892',
    'ch': 'CHE-123.456.788 TVA or CH TVA 123456',  # Swiss by Yannick Vaucher @ Camptocamp
    'cl': 'CL76086428-5',
    'co': 'CO213123432-1 or CO213.123.432-1',
    'cy': 'CY12345678F',
    'cz': 'CZ12345679',
    'de': 'DE123456788',
    'dk': 'DK12345674',
    'ee': 'EE123456780',
    'el': 'EL12345670',
    'es': 'ESA12345674',
    'fi': 'FI12345671',
    'fr': 'FR32123456789',
    'gb': 'GB123456782',
    'gr': 'GR12345670',
    'hu': 'HU12345676',
    'hr': 'HR01234567896',  # Croatia, contributed by Milan Tribuson
    'ie': 'IE1234567FA',
    'it': 'IT12345670017',
    'lt': 'LT123456715',
    'lu': 'LU12345613',
    'lv': 'LV41234567891',
    'mt': 'MT12345634',
    'mx': 'ABC123456T1B',
    'nl': 'NL123456782B90',
    'no': 'NO123456785',
    'pe': '10XXXXXXXXY or 20XXXXXXXXY or 15XXXXXXXXY or 16XXXXXXXXY or 17XXXXXXXXY',
    'pl': 'PL1234567883',
    'pt': 'PT123456789',
    'ro': 'RO1234567897',
    'se': 'SE123456789701',
    'si': 'SI12345679',
    'sk': 'SK0012345675',
    'tr': 'TR1234567890 (VERGINO) veya TR12345678901 (TCKIMLIKNO)'  # Levent Karakas @ Eska Yazilim A.S.
}


class ContactsXlsxDataWizard(models.TransientModel):
    _name = 'import.contacts.data.wizard'
    _description = 'Import Contacts xlsx data Wizard'

    xlsx_file = fields.Binary(string="Your xlsx file")

    def get_zip_in_format(self, zip):
        try:
            result = int(zip)
        except:
            result = zip

        return result

    def get_vat(self, vat):
        if vat:
            vat = str(vat).replace('.', '').replace(' ', '')
            if vat[:2].lower() not in _ref_vat.keys() or len(vat) != len(_ref_vat[vat[:2].lower()]):
                vat = ''
        return vat

    def import_contacts_xlsx_data(self):
        fp = tempfile.NamedTemporaryFile(suffix=".xlsx")
        fp.write(binascii.a2b_base64(self.xlsx_file))
        fp.seek(0)
        workbook = xlrd.open_workbook(fp.name)
        sheet = workbook.sheet_by_index(0)

        col_names = sheet.row_values(0)

        if len(col_names) == 21:
            self.import_db_contacts(sheet, col_names)
        elif len(col_names) == 15:
            self.import_clients_data(sheet, col_names)
        elif len(col_names) == 4:
            self.import_suppliers_data(sheet, col_names)

    def import_db_contacts(self, sheet, col_names):
        # Avoid persist data before the import is finished
        self._cr.autocommit(False)
        try:
            for i in range(1, sheet.nrows):
                row = sheet.row_values(i)

                partner_type = row[col_names.index("Type de partenaire")]
                name = row[col_names.index("Nom du partenaire")]
                street = row[col_names.index("Rue (destinataire facture)")]
                zip = row[col_names.index("CP (destin. facture)")]
                city = row[col_names.index("Ville destinataire de facture")]
                country_name = row[col_names.index("Pays (destin.facture)")]
                phone = row[col_names.index("Phone (destin.facture)")]
                title = row[col_names.index("Titre")]
                contact_name = row[col_names.index("Nom du contact")]
                job_position = row[col_names.index("Fonction")]
                tel = row[col_names.index("Téléphone 1")]
                cellphone = row[col_names.index("Téléphone portable")]
                fax = row[col_names.index("Fax")]
                email = row[col_names.index('E-mail')]
                contact_street = row[col_names.index('Adresse du contact')]
                contact_zip = row[col_names.index('Code Postal')]
                contact_city = row[col_names.index('Ville')]
                contact_country = row[col_names.index('Pays')]
                customer_rank = 1 if partner_type == 'Client' else 0
                supplier_rank = 1 if partner_type == 'Fournisseur' else 0

                partner = self.env['res.partner'].search([('name', '=', name)])
                if not partner:
                    partner = self.env['res.partner'].create({
                        'company_type': 'company',
                        'name': name,
                        'customer_rank': customer_rank,
                        'supplier_rank': supplier_rank,
                        'street': street if street != "x" else "",
                        'city': city if city != "x" else "",
                        'zip': self.get_zip_in_format(zip) if zip != "x" else "",
                        'country_id': self.env['res.country'].search([('name', '=', country_name)]).id,
                        'phone': phone,
                        'lang': 'fr_BE',
                    })
                else:
                    partner.write({
                        'customer_rank': partner.customer_rank if supplier_rank else customer_rank,
                        'supplier_rank': partner.supplier_rank if customer_rank else supplier_rank,
                    })

                title_exists = self.env['res.partner.title'].search([('name', '=', title)])
                if not title_exists:
                    title_exists = self.env['res.partner.title'].create({'name': title})
                contact = self.env['res.partner'].search([('name', '=', contact_name)])
                if not contact:
                    self.env['res.partner'].create({
                        'name': contact_name,
                        'company_type': 'person',
                        'title': title_exists.id,
                        'function': job_position,
                        'lang': 'fr_BE',
                        'phone': tel if tel != "xx" else "",
                        'mobile': cellphone if cellphone != "xx" else "",
                        'email': email if email != "x" else "",
                        'street': contact_street,
                        'city': contact_city,
                        'country_id': self.env['res.country'].search([('name', '=', contact_country)]).id,
                        'zip': self.get_zip_in_format(contact_zip),
                        'parent_id': partner.id
                    })

            self._cr.commit()

        except Exception as e:
            print(e)
            self._cr.rollback()
            raise Warning("Something goes wrong during the import!")

    def import_clients_data(self, sheet, col_names):
        self._cr.autocommit(False)
        try:
            for i in range(1, sheet.nrows):
                row = sheet.row_values(i)
                name = row[col_names.index("Partenaire")]
                code = row[col_names.index("Code Partenaire")]
                vat = row[col_names.index("Numéro d'identification d'entreprise")]
                correct_vat = self.get_vat(vat)
                print(correct_vat)
                partner = self.env['res.partner'].search([('name', '=', name)])
                if not partner:
                    self.env['res.partner'].create({
                        'name': name,
                        'customer_rank': 1,
                        'ref': code,
                        'vat': correct_vat,
                    })
                else:
                    partner_data = {'ref': code}
                    if not partner.vat:
                        partner_data['vat'] = correct_vat
                    partner.write(partner_data)

            self._cr.commit()

        except Exception as e:
            print(e)
            self._cr.rollback()
            raise Warning("Something goes wrong during the import!")

    def import_suppliers_data(self, sheet, col_names):
        self._cr.autocommit(False)
        try:
            for i in range(1, sheet.nrows):
                row = sheet.row_values(i)
                name = row[col_names.index("Partenaire")]
                code = row[col_names.index("Code Partenaire")]
                vat = row[col_names.index("Numéro d'identification d'entreprise")]
                correct_vat = self.get_vat(vat)
                iban = row[col_names.index("IBAN de la banque par défaut")]
                partner = self.env['res.partner'].search([('name', '=', name)])
                if not partner:
                    self.env['res.partner'].create({
                        'name': name,
                        'supplier_rank': 1,
                        'supplier_ref': code,
                        'vat': correct_vat,
                    })
                else:
                    partner_data = {'supplier_ref': code}
                    if not partner.vat:
                        partner_data['vat'] = correct_vat
                    partner.write(partner_data)

                if iban:
                    bank = self.env['res.partner.bank'].search([('acc_number', '=', iban)])
                    if not bank and partner.id:
                        self.env['res.partner.bank'].create({
                            'partner_id': partner.id,
                            'acc_number': iban,
                        })
            self._cr.commit()
        except Exception as e:
            print(e)
            self._cr.rollback()
            raise Warning("Something goes wrong during the import!")
