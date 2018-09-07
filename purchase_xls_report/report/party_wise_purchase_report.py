from odoo.addons.report_xlsx.report.report_xlsx import ReportXlsx
from odoo import fields, models,api

class PartyWisePurchaseReportXls(ReportXlsx):

    @api.multi
    def get_lines(self,partner_id,product_ids,date_from,date_to):
        
        lines = []
        for product_id in product_ids:
            
            purchase_obj = self.env['account.invoice.line'].search([
                                                           ('invoice_id.type','=','in_invoice'),
                                                           ('invoice_id.state', 'in', ['open','paid']),
                                                           ('invoice_id.partner_id', '=', partner_id.id),
                                                           ('invoice_id.journal_id.type', '=','purchase'),
                                                           ('invoice_id.date_invoice', '>=',date_from),
                                                           ('invoice_id.date_invoice', '<=',date_to),
                                                           ('product_id', '=', product_id.id),
                                                           
                                                           ])
            if purchase_obj:
                purchase_qty = 0.0
                purchase_amount = 0.0                                                  
                for purchase in purchase_obj:
                    purchase_qty += purchase.quantity
                    purchase_amount += purchase.price_subtotal
                 
                             
            #in case of purchase return   
                vals = {
                    'name': product_id.name,
                    'code':product_id.default_code,
                    'purchase_qty': purchase_qty,
                    'purchase_amount': purchase_amount,
                    'purchase_return_qty': 0.0,
                    'purchase_return_amount': 0.0,
                }
                lines.append(vals)
                


        return lines

    def generate_xlsx_report(self, workbook, data, lines):
        sheet = workbook.add_worksheet()
        format1 = workbook.add_format({'font_size': 14, 'bottom': True, 'right': True, 'left': True, 'top': True, 'align': 'vcenter', 'bold': True})
        format11 = workbook.add_format({'font_size': 12, 'align': 'center', 'right': True, 'left': True, 'bottom': True, 'top': True, 'bold': True})
        format21 = workbook.add_format({'font_size': 10, 'align': 'center', 'right': True, 'left': True,'bottom': True, 'top': True, 'bold': True})
        format21.set_num_format('#,##0.00')
        format3 = workbook.add_format({'bottom': True, 'top': True, 'font_size': 12})
        font_size_8 = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8})
        red_mark = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8,
                                        'bg_color': 'red'})
        justify = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 12})
        
        format3.set_align('center')
        font_size_8.set_align('center')
        justify.set_align('justify')
        format1.set_align('center')
        red_mark.set_align('center')
        sheet.merge_range(1, 1, 3, 7, 'Party Wise Purchase Report', format1)
        sheet.merge_range(4, 1, 5, 7, 'Period from :' +  data['form']['date_to'] + '  TO  ' + data['form']['date_from'], format1)

#         sheet.merge_range(4, 1, 5, 7, 'Period from : 01-Oct-2017 To 16-Nov-2017', format1)
#         
        sheet.write(7, 0,'Name / Code', format21)
        sheet.write(7, 1, 'Product Name', format21)
        sheet.write(7, 2, 'Description', format21)
        sheet.write(7, 3,'Avg Rate', format21)
        sheet.write(7, 4,'Qty purchase', format21)
        sheet.write(7, 5,'Amount purchase', format21)
#         sheet.write(7, 6,'Avg Freight', format21)
        sheet.write(7, 6,'Net purchase', format21)
    
        
        # report statrt
        product_row = 9        
        excel_col = 0
#         partner_ids = self.env['res.partner'].search([])
#         categ_ids = self.env['product.category'].search([])
        partner_ids = data['form']['partner']
        grand_purchase_sum =0.0
        grand_qty_sum=0.0
        
        for partner_id in partner_ids:
            partner_id =self.env['res.partner'].search([('id', '=', partner_id)])
            partner_name = partner_id.name
#             partner_name = partner_id.name
            sheet.write(product_row, 0, partner_name, format21)
            product_ids = self.env['product.product'].search([])#search([('categ_id', '=', categ_id.id)])
        
            get_lines = self.get_lines(partner_id,product_ids,data['form']['date_from'],data['form']['date_to'])
            if get_lines:
                sum_purchase_qty = 0.0
                sum_total_amount=0.0
                
                sum_return_qty =0.0
                sum_return_amount =0.0
            
                for each in get_lines:

                    sum_purchase_qty += int(each['purchase_qty'])
                    sum_total_amount += int(each['purchase_amount'])
                    
                    sum_return_qty += each['purchase_return_qty']
                    sum_return_amount += each['purchase_return_amount']
                                        
                    sheet.write(product_row + 1, excel_col + 0, each['code'], format21)
                    sheet.write(product_row + 1, excel_col + 1, each['name'], format21)
                    sheet.write(product_row + 1, excel_col + 2, each['name'], format21)
                    if each['purchase_qty'] != 0:
                        sheet.write(product_row + 1, excel_col + 3, (each['purchase_amount']/each['purchase_qty']), format21)
                    sheet.write(product_row + 1, excel_col + 4, each['purchase_qty'], format21)
                    sheet.write(product_row + 1, excel_col + 5, each['purchase_amount'], format21)
#                     sheet.write(product_row + 1, excel_col + 6, 0.0, format21)  
                    sheet.write(product_row + 1, excel_col + 6, each['purchase_amount'], format21)                   

                    product_row +=1 #lines adjestment of product
                

                product_row +=1 #lines adjestment of sum of qty and amount
                sheet.write(product_row, excel_col + 3, 'Sub Total')  
                sheet.write(product_row, excel_col + 4, sum_purchase_qty,format21)  
                sheet.write(product_row, excel_col + 5, sum_total_amount,format21) 
#                 sheet.write(product_row, excel_col + 6, 0.0)
                sheet.write(product_row, excel_col + 6, sum_total_amount,format21)  
                product_row +=1#lines adjestment of category heading 
                grand_purchase_sum +=sum_total_amount
                grand_qty_sum +=sum_purchase_qty
            
        sheet.write(product_row+1, excel_col + 1, "Grand Total")
        sheet.write(product_row+1, excel_col + 4, grand_qty_sum,format21)
        sheet.write(product_row+1, excel_col + 5, grand_purchase_sum,format21)
#         sheet.write(product_row+1, excel_col + 6, 0.0)
        sheet.write(product_row+1, excel_col + 6, grand_purchase_sum,format21)
        
        #for grand sum
        sum_purchase_per = 0.0
           
PartyWisePurchaseReportXls('report.export_purchaseinfo_xls.purchase_partywise_xls.xlsx', 'account.invoice')
