# -*- coding: utf-8 -*-
# from odoo import http


# class ProductExcelReportRynzn(http.Controller):
#     @http.route('/product_excel_report_rynzn/product_excel_report_rynzn/', auth='public')
#     def index(self, **kw):
#         return "Hello, world"

#     @http.route('/product_excel_report_rynzn/product_excel_report_rynzn/objects/', auth='public')
#     def list(self, **kw):
#         return http.request.render('product_excel_report_rynzn.listing', {
#             'root': '/product_excel_report_rynzn/product_excel_report_rynzn',
#             'objects': http.request.env['product_excel_report_rynzn.product_excel_report_rynzn'].search([]),
#         })

#     @http.route('/product_excel_report_rynzn/product_excel_report_rynzn/objects/<model("product_excel_report_rynzn.product_excel_report_rynzn"):obj>/', auth='public')
#     def object(self, obj, **kw):
#         return http.request.render('product_excel_report_rynzn.object', {
#             'object': obj
#         })
