from openerp import models, fields, api


class StockReport(models.TransientModel):
    _name = "wizard.purchase.history"
    _description = "Current Stock History"

    date_to= fields.Date("Date To")
    date_from= fields.Date("Date From")
    report_type = fields.Selection([('indivproduct_wise','Product Wise Report'),('product_wise','Product Cateogory Wise Report'),('grand_summary','Grand Summary'),('purchase_partywise','Party Wise')],string='Relative')
    category = fields.Many2many('product.category',  string='Categories')
    warehouse = fields.Many2many('stock.warehouse', string='Warehouse')

    partner = fields.Many2many('res.partner',  string='Partner')
    indv_product = fields.Many2many('product.product',  string='Products')


    @api.multi
    def export_xls(self):
        context = self._context
        datas = {'ids': context.get('active_ids', [])}
        datas['model'] = 'purchase.report'
        datas['form'] = self.read()[0]
        for field in datas['form'].keys():
            if isinstance(datas['form'][field], tuple):
                datas['form'][field] = datas['form'][field][0]
        if context.get('xls_export'):
            if datas['form']['report_type'] == 'grand_summary':
                return {'type': 'ir.actions.report.xml',
                        'report_name': 'export_purchaseinfo_xls.summr_wise_xls.xlsx',
                        'datas': datas,
                        'name': 'purchase'
                        }
            elif datas['form']['report_type'] == 'indivproduct_wise':
                return {'type': 'ir.actions.report.xml',
                        'report_name': 'export_purchaseinfo_xls.purchase_indivproductwise_xls.xlsx',
                        'datas': datas,
                        'name': 'purchase'
                        }
            elif datas['form']['report_type'] == 'product_wise':
                return {'type': 'ir.actions.report.xml',
                        'report_name': 'export_purchaseinfo_xls.purchase_productwise_xls.xlsx',
                        'datas': datas,
                        'name': 'purchase'
                        }
            
            elif datas['form']['report_type'] == 'purchase_partywise':
                return {'type': 'ir.actions.report.xml',
                        'report_name': 'export_purchaseinfo_xls.purchase_partywise_xls.xlsx',
                        'datas': datas,
                        'name': 'purchase'
                        }
            
#             elif datas['form']['report_type'] == 'categ_wise':
#                 return {'type': 'ir.actions.report.xml',
#                         'report_name': 'export_purchaseinfo_xls.categ_wise_xls.xlsx',
#                         'datas': datas,
#                         'name': 'purchase'
#                         }