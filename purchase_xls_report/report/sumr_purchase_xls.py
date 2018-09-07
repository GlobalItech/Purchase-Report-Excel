from odoo.addons.report_xlsx.report.report_xlsx import ReportXlsx
from odoo import fields, models,api

class purchaseSummReportXls(ReportXlsx):

    @api.multi
    def get_lines(self,product_ids,date_from,date_to):
        
        lines = []
            
        for product_id in product_ids:
#             date_from =str(date_from)
#             date_to= str(date_to)
            purchase_obj = self.env['account.invoice.line'].search([
                                                            ('invoice_id.type','=','in_invoice'),
                                                            ('invoice_id.state', 'in', ['open','paid']),
                                                           ('invoice_id.date_invoice', '>=',date_from),
                                                           ('invoice_id.date_invoice', '<=',date_to),
                                                           ('invoice_id.journal_id.type', '=','purchase'),
                                                            ('invoice_id.origin', 'ilike', 'PO'),
                                                           ('product_id', '=', product_id.id)])
            if purchase_obj:
                purchase_qty = 0.0
                purchase_amount = 0.0
                for purchase in purchase_obj:
                    purchase_qty += purchase.quantity
                    purchase_amount += purchase.price_subtotal
                purchase_return_obj = self.env['account.invoice.line'].search([
                                    ('invoice_id.state', 'in', ['open', 'paid']),
                                    ('invoice_id.date_invoice', '>=', date_from),
                                    ('invoice_id.date_invoice', '<=', date_to),
                                    ('invoice_id.journal_id.type', '=', 'purchase'),
                                    ('invoice_id.origin', 'ilike','BILL' ),
                                    ('product_id', '=', product_id.id)])
                if purchase_return_obj:
                    return_qty = 0.0
                    return_amount = 0.0
                    for purchase_return in purchase_return_obj:
                        return_qty += purchase_return.quantity
                        return_amount += purchase_return.price_subtotal

                    vals = {
                        'name': product_id.name,
                        'purchase_qty': purchase_qty,
                        'purchase_amount': purchase_amount,
                        'purchase_return_qty': return_qty,
                        'purchase_return_amount': return_amount,
                    }
                    lines.append(vals)
                else:
                    vals = {
                        'name': product_id.name,
                        'purchase_qty': purchase_qty,
                        'purchase_amount': purchase_amount,
                        'purchase_return_qty': 0.0,
                        'purchase_return_amount': 0.0,
                    }
                    lines.append(vals)

            # if purchase_obj:
            #     purchase_qty = 0.0
            #     purchase_amount = 0.0
            #     for purchase in purchase_obj:
            #         purchase_qty += purchase.quantity
            #         purchase_amount += purchase.price_subtotal
            #     vals = {
            #         'name': product_id.name,
            #         'purchase_qty': purchase_qty,
            #         'purchase_amount': purchase_amount,
            #         'purchase_return_qty': 0.0,
            #         'purchase_return_amount': 0.0,
            #     }
            #     lines.append(vals)
                 
                             
            #in case of purchase return


                # purchase_return_obj = self.env['stock.pack.operation'].search([('state', '=', 'done'),
                #                                                             ('product_id', '=', product_id.id),
                #                                                             ('date', '>=', date_from),
                #                                                             ('date', '<=', date_to),
                #                                                             ('location_id.usage', '=', 'customer'),])

            # if purchase_obj:
            #     purchase_qty = 0.0
            #     purchase_amount = 0.0
            #     for purchase in purchase_obj:
            #         purchase_qty += purchase.quantity
            #         purchase_amount += purchase.price_subtotal
            #     if not purchase_return_obj:
            #         vals = {
            #             'name': product_id.name,
            #             'purchase_qty': purchase_qty,
            #             'purchase_amount': purchase_amount,
            #             'purchase_return_qty': 0.0,
            #             'purchase_return_amount': 0.0,
            #         }
            #         lines.append(vals)
            # if purchase_return_obj:
            #     return_qty = 0.0
            #     return_amount = 0.0
            #     for purchase_return in purchase_return_obj:
            #         return_qty += purchase_return.quantity
            #         return_amount += purchase_return.price_subtotal
            #         # return_qty += purchase_return.qty_done
            #         # return_amount += return_qty * purchase_return.product_id.lst_price
            #     if not purchase_obj:
            #         vals = {
            #             'name': product_id.name,
            #             'purchase_qty': 0.0,
            #             'purchase_amount': 0.0,
            #             'purchase_return_qty': return_qty,
            #             'purchase_return_amount': return_amount,
            #         }
            #     lines.append(vals)
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
#         style = workbook.add_format('align: wrap yes; borders: top thin, bottom thin, left thin, right thin;')
#         style.num_format_str = '#,##0.00'
        format3.set_align('center')
        font_size_8.set_align('center')
        justify.set_align('justify')
        format1.set_align('center')
        red_mark.set_align('center')
       
        sheet.merge_range(1, 1, 3, 7, 'Purchase Summary', format1)
        sheet.merge_range(4, 1, 5, 7, 'Period from :' +  data['form']['date_from'] + '  TO  ' + data['form']['date_to'], format1)

#         sheet.merge_range(4, 1, 5, 7, 'Period from : 01-Oct-2017 To 16-Nov-2017', format1)
#         
        sheet.merge_range(7, 0, 8, 0,'Description', format21)
        sheet.merge_range(7, 1, 7, 2, 'Gross purchase', format21)
        sheet.merge_range(7, 3, 7, 4, 'Return purchase', format21)
        sheet.merge_range(7, 5, 7, 6, 'Net purchase', format21)
#         sheet.merge_range(7, 7, 8, 7,'Freight', format21)
        sheet.merge_range(7, 7, 8, 7,'Grand Net purchase', format21)
        sheet.merge_range(7, 8, 8, 8,'% Purchase', format21)
        sheet.merge_range(7, 9, 8, 9,'Avg Rate on Net purchase', format21)
         
        
        sheet.write(8, 1, 'QTY', format21)
        sheet.write(8, 2, 'Amount', format21)
        sheet.write(8, 3, 'QTY', format21)
        sheet.write(8, 4, 'Amount', format21)
        sheet.write(8, 5, 'QTY', format21)
        sheet.write(8, 6, 'Amount', format21)
        
        
        
        # report statrt
        product_row = 9        
        excel_col = 0
#         categ_ids = self.env['product.category'].search([])
        if data['form']['category']:
            categ_ids = data['form']['category']
        else:
            categ_ids = self.env['product.category'].search([])
        grand_sum =0.0
        total_excel_rows =[]
        for categ_id in categ_ids:
            product_ids = self.env['product.product'].search([('categ_id', '=', categ_id.id)])
            ctag_name =self.env['product.category'].search([('id', '=', categ_id.id)]).name
            
            get_lines = self.get_lines(product_ids,data['form']['date_from'],data['form']['date_to'])
            if get_lines:
                sum_purchase_qty = 0.0
                sum_total_amount=0.0
                
                sum_return_qty =0.0
                sum_return_amount =0.0
                
                sum_net_purchase =0.0
                sum_net_amount = 0.0
                
                sum_avg_net_purchase=0.0
                categ_name = ctag_name
                sheet.write(product_row, 0, categ_name, format21)
                for each in get_lines:
                    values ={
                            'row_no':product_row +1,
                            'row_grand_value': (each['purchase_amount'] -  each['purchase_return_amount']),
                            }
                    total_excel_rows.append(values)
#                     total_excel_rows.append(product_row +1)
                    sum_purchase_qty += int(each['purchase_qty'])
                    sum_total_amount += int(each['purchase_amount'])
                    
                    # sum_return_qty += each['purchase_return_qty']
                    # sum_return_amount += each['purchase_return_amount']
                    
                    sheet.write(product_row + 1, excel_col + 0, each['name'], format21)
                    sheet.write(product_row + 1, excel_col + 1, each['purchase_qty'], format21)
                    sheet.write(product_row + 1, excel_col + 2, each['purchase_amount'], format21)                    
                    #purchase return
                    sum_return_qty += each['purchase_return_qty']
                    sum_return_amount += each['purchase_return_amount']
                    sheet.write(product_row + 1, excel_col + 3, each['purchase_return_qty'], format21)
                    sheet.write(product_row + 1, excel_col + 4, each['purchase_return_amount'], format21)
                    
                    # Net purchase qty and Amount
                    sum_net_purchase +=(each['purchase_qty'] - each['purchase_return_qty'] )
                    sum_net_amount += (each['purchase_amount'] -  each['purchase_return_amount']) 
                    
                    
                    net_purchase = (each['purchase_qty'] - each['purchase_return_qty'] )
                    net_amount =(each['purchase_amount'] -  each['purchase_return_amount'])
                    if net_purchase !=0:
                        sum_avg_net_purchase += (net_amount/net_purchase)
                    
                    sheet.write(product_row + 1, excel_col + 5, net_purchase , format21)
                    sheet.write(product_row + 1, excel_col + 6, net_amount , format21)
                    
                    # % purchase and avg rate on net purchase
#                     sheet.write(product_row + 1, excel_col + 7, 0.0, format21)
                    #grand net purchase
                    sheet.write(product_row + 1, excel_col + 7, (each['purchase_amount'] - each['purchase_return_amount']) , format21)
#                     sheet.write(product_row + 1, excel_col + 9, 0.0 , format21)
                    if net_purchase !=0:
                        sheet.write(product_row + 1, excel_col + 9, (net_amount/net_purchase), format21)
                    product_row +=1 #lines adjestment of product
                

                product_row +=1 #lines adjestment of sum of qty and amount
                sheet.write(product_row, excel_col + 1, sum_purchase_qty,format21)
                sheet.write(product_row, excel_col + 2, sum_total_amount,format21)
                 
                sheet.write(product_row, excel_col + 3, sum_return_qty,format21)
                sheet.write(product_row, excel_col + 4, sum_return_amount,format21)
                
                sheet.write(product_row, excel_col + 5, sum_net_purchase,format21)
                sheet.write(product_row, excel_col + 6, sum_net_amount,format21)
#                 sheet.write(product_row, excel_col + 7, 0, format21)
                #grand net total
                sheet.write(product_row, excel_col + 7, sum_net_amount,format21)
                sheet.write(product_row, excel_col + 9, sum_avg_net_purchase,format21)
                product_row +=1#lines adjestment of category heading 
                grand_sum +=sum_net_amount
            
        sheet.write(product_row+1, excel_col + 8, "Grand Total")
        sheet.write(product_row+1, excel_col + 6, grand_sum,format21)
        sheet.write(product_row+1, excel_col + 8, grand_sum,format21)
        
        
        #for grand sum
        sum_purchase_per = 0.0
        for total_excel_row in total_excel_rows:
            if grand_sum !=0:
                sum_purchase_per += (total_excel_row['row_grand_value']/grand_sum)*100
            if grand_sum == 0:
                sheet.write(total_excel_row['row_no'], 8, 0,format21)
            else:
                sheet.write(total_excel_row['row_no'], 8, (total_excel_row['row_grand_value']/grand_sum)*100,format21)
        sheet.write(product_row+1, excel_col + 8, sum_purchase_per,format21)
            
purchaseSummReportXls('report.export_purchaseinfo_xls.summr_wise_xls.xlsx', 'account.invoice')
