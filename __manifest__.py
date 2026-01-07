# -*- coding: utf-8 -*-
{
    'name': "Reporte de asistencias",

    'summary': """
        Módulo para generar reportes de asistencias
        """,

    'description': """
        Módulo personalizado para generar reportes de asistencias en Odoo. Permite a los usuarios crear y visualizar informes detallados sobre las asistencias registradas en el sistema.
    """,

    'author': "GonzaOdoo",
    'website': "http://www.yourcompany.com",

    # Categories can be used to filter modules in modules listing
    # Check https://github.com/odoo/odoo/blob/master/odoo/addons/base/module/module_data.xml
    # for the full list
    'category': 'Uncategorized',
    'version': '1.0',

    # any module necessary for this one to work correctly
    'depends': ['hr_attendance','hr_payroll'],

    # always loaded
    "data": ["security/ir.model.access.csv",
             "views/report_assistance.xml",
            ],
}
