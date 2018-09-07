{
    'name': 'Export Purchase Report in Excel',
    'version': '0.2',
    'category': 'Warehouse',
    'license': "AGPL-3",
    'summary': "purchase",
    'author': 'Itech reosurces',
    'company': 'ItechResources',
    'depends': [
                'account',
                'stock',
                'purchase',
                'report_xlsx'
                ],
    'data': [
            'views/wizard_view.xml',
#             'views/territory_form.xml',
#             'views/customer_sale_recovery.xml',
            ],
    'images': ['static/description/banner.jpg'],
    'installable': True,
    'auto_install': False,
    "license": 'OPL-1',
    'price':'20.0',
    'currency': 'USD',
}
