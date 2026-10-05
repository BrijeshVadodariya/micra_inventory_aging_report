# -*- coding: utf-8 -*-

from odoo import models, fields, api


class PurchaseOrderLine(models.Model):
    _inherit = 'purchase.order.line'

    previous_price = fields.Float(
        string="Previous Price",
        compute='_compute_previous_price',
        store=True,
        readonly=True,
        digits='Product Price',
        help="Price of the last purchase for this product."
    )
    growth = fields.Float(
        string="Growth",
        compute='_compute_growth',
        store=True,
        readonly=True,
        digits=(16, 2),
        help="Percentage price growth compared to previous purchase price."
    )
    growth_display = fields.Char(
        string="Growth (%)",
        compute='_compute_growth',
        store=True,
        readonly=True,
        help="Formatted percentage price growth with sign."
    )

    @api.depends('product_id', 'order_id.date_order', 'order_id.partner_id', 'company_id')
    def _compute_previous_price(self):
        for line in self:
            if not line.product_id:
                line.previous_price = 0.0
                continue

            order_origin_id = line.order_id._origin.id if line.order_id else False
            line_origin_id = line._origin.id if line.id else False
            comp_id = line.company_id.id or self.env.company.id

            # 1. Look for the most recent confirmed purchase order line for the same product
            domain = [
                ('product_id', '=', line.product_id.id),
                ('state', 'in', ['purchase', 'done']),
                ('company_id', '=', comp_id),
            ]
            if order_origin_id:
                domain.append(('order_id', '!=', order_origin_id))
            if line_origin_id:
                domain.append(('id', '!=', line_origin_id))

            last_line = self.search(domain, order='date_order desc, id desc', limit=1)

            # 2. Fallback: look for any previous purchase order line (draft/sent) if none confirmed
            if not last_line:
                domain_any = [
                    ('product_id', '=', line.product_id.id),
                    ('state', '!=', 'cancel'),
                    ('company_id', '=', comp_id),
                ]
                if order_origin_id:
                    domain_any.append(('order_id', '!=', order_origin_id))
                if line_origin_id:
                    domain_any.append(('id', '!=', line_origin_id))
                last_line = self.search(domain_any, order='date_order desc, id desc', limit=1)

            if last_line:
                line.previous_price = last_line.price_unit
            else:
                line.previous_price = line.product_id.standard_price or 0.0

    @api.depends('price_unit', 'previous_price')
    def _compute_growth(self):
        for line in self:
            prev = line.previous_price or 0.0
            curr = line.price_unit or 0.0

            if prev > 0:
                diff = curr - prev
                rate = round((diff / prev) * 100.0, 2)
                line.growth = rate
                if rate > 0:
                    line.growth_display = f"+{rate:.2f}%"
                elif rate < 0:
                    line.growth_display = f"{rate:.2f}%"
                else:
                    line.growth_display = "0.00%"
            elif prev == 0 and curr > 0:
                line.growth = 100.0
                line.growth_display = "+100.00%"
            else:
                line.growth = 0.0
                line.growth_display = "0.00%"

    @api.onchange('product_id')
    def _onchange_product_id_previous_price(self):
        if self.product_id:
            order_origin_id = self.order_id._origin.id if self.order_id else False
            comp_id = self.company_id.id or self.env.company.id

            domain = [
                ('product_id', '=', self.product_id.id),
                ('state', 'in', ['purchase', 'done']),
                ('company_id', '=', comp_id),
            ]
            if order_origin_id:
                domain.append(('order_id', '!=', order_origin_id))

            last_line = self.env['purchase.order.line'].search(domain, order='date_order desc, id desc', limit=1)
            if not last_line:
                domain_any = [
                    ('product_id', '=', self.product_id.id),
                    ('state', '!=', 'cancel'),
                    ('company_id', '=', comp_id),
                ]
                if order_origin_id:
                    domain_any.append(('order_id', '!=', order_origin_id))
                last_line = self.env['purchase.order.line'].search(domain_any, order='date_order desc, id desc', limit=1)

            self.previous_price = last_line.price_unit if last_line else (self.product_id.standard_price or 0.0)
            self._compute_growth()

    @api.onchange('price_unit', 'previous_price')
    def _onchange_price_unit_growth(self):
        self._compute_growth()
