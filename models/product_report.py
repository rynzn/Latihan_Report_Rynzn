import base64
import xlsxwriter

from odoo import models, fields, api, _
try:
    from StringIO import StringIO
except ImportError:
    from io import BytesIO

class ProductReportWizard(models.TransientModel):
    _name = "product.report.wizard"
    _description = 'Product Report Wizard'

    name = fields.Char(string='Name', default="PRODUCT LIST")
    product_type = fields.Selection(
        [("all","All"),
        ("service","Service"),
        ("product","Product"),
        ("consu","Consumanble")],
        string='Product Type')
    sale_ok = fields.Boolean(string='Can be Sold')
    purchase_ok = fields.Boolean(string='Can be Purchase')
    date_created = fields.Date(string='Date Created', default=fields.Date.today())
    filename = fields.Char(string='Filename')
    data_file = fields.Binary(string='Data file')

    def export(self):

        # Membuat Worksheet
        folder_title = self.name + "-" + str(self.date_created) + ".xlsx"
        file_data = BytesIO()
        workbook = xlsxwriter.Workbook(file_data)
        ws = workbook.add_worksheet((self.name))  

        # Menambahkan style
        style = workbook.add_format({'left': 1, 'top': 1,'right':1,'bold': True,'fg_color': '#339966','font_color': 'white','align':'center'})
        style.set_text_wrap()
        style.set_align('vcenter')
        style_bold = workbook.add_format({'left': 1, 'top': 1,'right':1,'bottom':1,'bold': True,'align':'center','num_format':'_(Rp* #,##0_);_(Rp* (#,##0);_(* "-"??_);_(@_)'})
        style_bold_orange = workbook.add_format({'left': 1, 'top': 1,'right':1,'bold': True,'align':'center','fg_color': '#FF6600','font_color': 'white'})
        style_no_bold = workbook.add_format({'left': 1,'right':1,'bottom':1, 'num_format':'_(Rp* #,##0_);_(Rp* (#,##0);_(* "-"??_);_(@_)'})

        # Mengumpulkan dan filter data
        if self.product_type == 'all':
            product_line = self.env['product.product'].search([
                ('sale_ok','=', self.sale_ok),
                ('purchase_ok','=',self.purchase_ok),])
        else:
            product_line = self.env['product.product'].search([
                ('sale_ok','=', self.sale_ok),
                ('purchase_ok','=',self.purchase_ok),
                ('type','=',self.product_type),])

        # Membuat Column Header
        ws.merge_range('A1:D1',  self.name + ' ' + str(self.date_created), style_bold)
        ws.set_column(1, 1, 10)
        ws.set_column(1, 2, 40)
        ws.set_column(1, 3, 25)
        ws.set_column(1, 4, 25)
        ws.write(3, 0,'NO ', style_bold_orange)
        ws.write(3, 1,'NAMA ', style_bold_orange)
        ws.write(3, 2, 'TIPE', style_bold_orange)
        ws.write(3, 3, 'HARGA ', style_bold_orange)

        row_count = 4
        count = 1

        # Mengisi data pada setiap baris & kolom
        for product in product_line:
            ws.write(row_count, 0, str(count), style_no_bold)
            ws.write(row_count, 1,product.name, style_no_bold)
            ws.write(row_count, 2,product.type, style_no_bold)
            ws.write(row_count, 3,product.price, style_no_bold)
            count+=1
            row_count+=1

        # Menyimpan data di field data_file
        workbook.close()        
        out = base64.encodestring(file_data.getvalue())
        self.write({'data_file': out, 'filename': folder_title})

        return self.view_form()

    def view_form(self):        
        view = self.env.ref('product_excel_report_rynzn.view_product_excel_report_wizard_form')
        return {
            'name': _('Product Report Wizard'),
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'product.report.wizard',
            'views': [(view.id, 'form')],
            'res_id': self.id,
            'type': 'ir.actions.act_window',
            'target': 'new',
        }