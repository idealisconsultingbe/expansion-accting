# -*- coding: utf-8 -*-
##############################################################################
#
# This module is developed by Idealis Consulting SPRL
# Copyright (C) 2019 Idealis Consulting SPRL (<https://idealisconsulting.com>).
# All Rights Reserved
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################

{
    'name': 'EXP Contacts Import',
    'summary': 'Module for importing contacts (Customers and Suppliers) data',
    'version': '0.1',
    'description': """
Manage the possibility to import Contacts data in Odoo
        """,
    'author': 'rdb@idealisconsulting.com - Idealis Consulting',
    'depends': [
        'account_accountant',
        'exp_partners',
    ],
    'data': [
        'views/import_wizard_view.xml',
    ],
    'installable': True,
    'auto_install': False,
}
