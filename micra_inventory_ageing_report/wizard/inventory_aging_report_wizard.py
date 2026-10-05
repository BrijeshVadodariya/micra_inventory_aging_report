# -*- coding: utf-8 -*-

from datetime import datetime, time, timedelta
import base64
import io
import logging

from odoo import models, fields, api, _
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


class InventoryAgingReportWizard(models.TransientModel):
    _name = 'inventory.aging.report.wizard'
    _description = 'Inventory Aging Report Wizard'

    aging_date = fields.Date(
        string="Aging Date",
        default=lambda self: fields.Date.context_today(self) - timedelta(days=30),
        required=True,
        help="Aging date to calculate the age of stock from today back to this date."
    )
    period_length = fields.Integer(
        string="Period Length (Days)",
        default=30,
        required=True,
        help="Number of days in each aging bracket (e.g., 30 days)."
    )
    company_id = fields.Many2one(
        'res.company',
        string="Company",
        default=lambda self: self.env.company,
        required=True,
        help="Select company for the report."
    )
    warehouse_ids = fields.Many2many(
        'stock.warehouse',
        string="Warehouses",
        help="Select warehouses to filter stock data."
    )
    location_ids = fields.Many2many(
        'stock.location',
        string="Internal Locations",
        domain=[('usage', '=', 'internal')],
        help="Select internal stock locations to filter stock data."
    )
    filter_by = fields.Selection(
        [
            ('product', 'Product'),
            ('category', 'Product Category'),
        ],
        string="Filter By",
        default='product',
        required=True,
        help="Filter report data by specific products or product categories."
    )
    product_ids = fields.Many2many(
        'product.product',
        string="Products",
        help="Select products for the report."
    )
    category_ids = fields.Many2many(
        'product.category',
        string="Product Category",
        help="Select product categories for the report."
    )
    line_ids = fields.One2many(
        'inventory.aging.report.line',
        'wizard_id',
        string="Report Lines"
    )

    @api.constrains('period_length')
    def _check_period_length(self):
        for record in self:
            if record.period_length <= 0:
                raise UserError(_("Period length must be greater than 0."))

    @api.model
    def _compute_bracket_labels_dict(self, p):
        if not p or p <= 0:
            p = 30
        if p == 1:
            return {
                'b1': "0 Days (Today)",
                'b2': "1 Day Ago",
                'b3': "2 Days Ago",
                'b4': "3 Days Ago",
                'b5': "4+ Days Ago",
            }
        elif p < 30:
            return {
                'b1': f"0-{p - 1} Days",
                'b2': f"{p}-{2 * p - 1} Days",
                'b3': f"{2 * p}-{3 * p - 1} Days",
                'b4': f"{3 * p}-{4 * p - 1} Days",
                'b5': f"{4 * p}+ Days",
            }
        return {
            'b1': f"0-{p}",
            'b2': f"{p + 1}-{2 * p}",
            'b3': f"{2 * p + 1}-{3 * p}",
            'b4': f"{3 * p + 1}-{4 * p}",
            'b5': f"{4 * p + 1}+",
        }

    def _get_bracket_labels(self):
        """Returns labels for the 5 aging brackets based on period length."""
        self.ensure_one()
        return self._compute_bracket_labels_dict(self.period_length)

    def _get_target_locations(self):
        """Retrieve target internal stock locations including child locations."""
        self.ensure_one()
        Location = self.env['stock.location']
        if self.location_ids:
            return Location.search([
                ('usage', '=', 'internal'),
                ('id', 'child_of', self.location_ids.ids),
            ])
        elif self.warehouse_ids:
            view_locations = self.warehouse_ids.mapped('view_location_id')
            return Location.search([
                ('usage', '=', 'internal'),
                ('id', 'child_of', view_locations.ids),
            ])
        else:
            return Location.search([
                ('usage', '=', 'internal'),
                ('company_id', 'in', [False, self.company_id.id]),
            ])

    def _get_aging_data(self):
        """Calculates FIFO stock aging for all products with on-hand inventory."""
        self.ensure_one()
        locations = self._get_target_locations()
        if not locations:
            return []

        company_id = self.company_id.id
        today = fields.Date.context_today(self)
        start_date = self.aging_date or (today - timedelta(days=30))
        start_dt = datetime.combine(start_date, time.min)
        end_dt = datetime.combine(today, time.max)
        p = self.period_length or 30

        # Build location to warehouse mapping
        warehouses = self.env['stock.warehouse'].search([
            ('company_id', 'in', [False, company_id])
        ])
        loc_to_warehouse = {}
        for wh in warehouses:
            wh_locs = self.env['stock.location'].search([
                ('usage', '=', 'internal'),
                ('id', 'child_of', wh.view_location_id.id),
            ])
            for loc in wh_locs:
                loc_to_warehouse[loc.id] = wh

        default_warehouse = warehouses[:1] if warehouses else False

        # 1. Query on-hand stock quants in target internal locations
        quant_domain = [
            ('location_id', 'in', locations.ids),
            ('company_id', 'in', [False, company_id]),
            ('quantity', '>', 0),
        ]
        if self.filter_by == 'product' and self.product_ids:
            quant_domain.append(('product_id', 'in', self.product_ids.ids))
        elif self.filter_by == 'category' and self.category_ids:
            matched_prods = self.env['product.product'].search([
                ('categ_id', 'child_of', self.category_ids.ids),
                ('company_id', 'in', [False, company_id]),
            ])
            if not matched_prods:
                return []
            quant_domain.append(('product_id', 'in', matched_prods.ids))

        quants = self.env['stock.quant'].search(quant_domain)

        # Aggregate on-hand stock per (product_id, warehouse_id)
        stock_on_hand = {}
        for q in quants:
            product = q.product_id
            wh = loc_to_warehouse.get(q.location_id.id, default_warehouse)
            wh_id = wh.id if wh else False
            key = (product.id, wh_id)
            if key not in stock_on_hand:
                stock_on_hand[key] = {
                    'product': product,
                    'warehouse': wh,
                    'qty': 0.0,
                }
            stock_on_hand[key]['qty'] += q.quantity

        if not stock_on_hand:
            return []

        # 2. For each product/warehouse on-hand stock, allocate using FIFO incoming moves
        report_data = []

        for (prod_id, wh_id), stock_info in stock_on_hand.items():
            product = stock_info['product']
            warehouse = stock_info['warehouse']
            total_qty_on_hand = stock_info['qty']

            if total_qty_on_hand <= 0:
                continue

            unit_cost = product.standard_price or 0.0

            if warehouse:
                wh_locations = [loc_id for loc_id, w in loc_to_warehouse.items() if w.id == warehouse.id]
            else:
                wh_locations = locations.ids

            # Find past incoming stock moves for this product up to today (newest first)
            move_domain = [
                ('product_id', '=', prod_id),
                ('state', '=', 'done'),
                ('location_dest_id', 'in', wh_locations if wh_locations else locations.ids),
                ('location_id', 'not in', locations.ids),
                ('date', '<=', end_dt),
                ('company_id', 'in', [False, company_id]),
            ]
            incoming_moves = self.env['stock.move'].search(move_domain, order='date desc, id desc')

            b_qty = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.0}
            b_val = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.0}

            remaining_qty = total_qty_on_hand

            for move in incoming_moves:
                if remaining_qty <= 0:
                    break

                move_qty = move.product_qty or getattr(move, 'quantity', 0.0) or move.product_uom_qty or 0.0
                if move_qty <= 0:
                    continue

                allocated_qty = min(move_qty, remaining_qty)

                move_dt = move.date
                move_date = move_dt.date() if isinstance(move_dt, datetime) else move_dt
                age_days = (today - move_date).days
                if age_days < 0:
                    age_days = 0

                # Determine unit price for this incoming move
                m_price = 0.0
                if hasattr(move, 'purchase_line_id') and move.purchase_line_id and move.purchase_line_id.price_unit:
                    m_price = move.purchase_line_id.price_unit
                elif getattr(move, 'price_unit', 0.0):
                    m_price = move.price_unit
                elif hasattr(move, 'stock_valuation_layer_ids') and move.stock_valuation_layer_ids:
                    m_price = move.stock_valuation_layer_ids[0].unit_cost
                if not m_price or m_price <= 0:
                    m_price = unit_cost

                # Determine aging bracket
                if age_days < p:
                    bucket = 1
                elif age_days < 2 * p:
                    bucket = 2
                elif age_days < 3 * p:
                    bucket = 3
                elif age_days < 4 * p:
                    bucket = 4
                else:
                    bucket = 5

                b_qty[bucket] += allocated_qty
                b_val[bucket] += allocated_qty * m_price
                remaining_qty -= allocated_qty

            # If remaining on-hand quantity has no recent move, allocate to oldest bracket (Bucket 5)
            if remaining_qty > 0:
                b_qty[5] += remaining_qty
                b_val[5] += remaining_qty * unit_cost

            # Calculate weighted average price for each bracket
            b1_price = round(b_val[1] / b_qty[1], 2) if b_qty[1] > 0 else 0.0
            b2_price = round(b_val[2] / b_qty[2], 2) if b_qty[2] > 0 else 0.0
            b3_price = round(b_val[3] / b_qty[3], 2) if b_qty[3] > 0 else 0.0
            b4_price = round(b_val[4] / b_qty[4], 2) if b_qty[4] > 0 else 0.0
            b5_price = round(b_val[5] / b_qty[5], 2) if b_qty[5] > 0 else 0.0

            tot_qty = round(sum(b_qty.values()), 2)
            tot_val = round(sum(b_val.values()), 2)

            report_data.append({
                'product_id': product.id,
                'default_code': product.default_code or '',
                'product_name': product.display_name or product.name or '',
                'categ_id': product.categ_id.id if product.categ_id else False,
                'categ_name': product.categ_id.display_name if product.categ_id and product.categ_id.display_name else 'Uncategorized',
                'warehouse_id': warehouse.id if warehouse else False,
                'warehouse_name': warehouse.name if warehouse and warehouse.name else 'N/A',
                'unit_cost': unit_cost,
                'bracket_1_price': b1_price,
                'bracket_1_qty': round(b_qty[1], 2),
                'bracket_1_val': round(b_val[1], 2),
                'bracket_2_price': b2_price,
                'bracket_2_qty': round(b_qty[2], 2),
                'bracket_2_val': round(b_val[2], 2),
                'bracket_3_price': b3_price,
                'bracket_3_qty': round(b_qty[3], 2),
                'bracket_3_val': round(b_val[3], 2),
                'bracket_4_price': b4_price,
                'bracket_4_qty': round(b_qty[4], 2),
                'bracket_4_val': round(b_val[4], 2),
                'bracket_5_price': b5_price,
                'bracket_5_qty': round(b_qty[5], 2),
                'bracket_5_val': round(b_val[5], 2),
                'total_qty': tot_qty,
                'total_val': tot_val,
            })

        report_data.sort(key=lambda x: (str(x['categ_name'] or ''), str(x['product_name'] or '')))
        return report_data

    def _get_aging_data_grouped(self):
        """Returns aging data structured and grouped by product category for reports."""
        self.ensure_one()
        raw_data = self._get_aging_data()
        grouped = {}
        grand_total = {
            'bracket_1_price': 0.0, 'bracket_1_qty': 0.0, 'bracket_1_val': 0.0,
            'bracket_2_price': 0.0, 'bracket_2_qty': 0.0, 'bracket_2_val': 0.0,
            'bracket_3_price': 0.0, 'bracket_3_qty': 0.0, 'bracket_3_val': 0.0,
            'bracket_4_price': 0.0, 'bracket_4_qty': 0.0, 'bracket_4_val': 0.0,
            'bracket_5_price': 0.0, 'bracket_5_qty': 0.0, 'bracket_5_val': 0.0,
            'total_qty': 0.0, 'total_val': 0.0,
        }

        for item in raw_data:
            categ = item['categ_name']
            if categ not in grouped:
                grouped[categ] = {
                    'lines': [],
                    'total': {
                        'unit_cost': 0.0,
                        'bracket_1_price': 0.0, 'bracket_1_qty': 0.0, 'bracket_1_val': 0.0,
                        'bracket_2_price': 0.0, 'bracket_2_qty': 0.0, 'bracket_2_val': 0.0,
                        'bracket_3_price': 0.0, 'bracket_3_qty': 0.0, 'bracket_3_val': 0.0,
                        'bracket_4_price': 0.0, 'bracket_4_qty': 0.0, 'bracket_4_val': 0.0,
                        'bracket_5_price': 0.0, 'bracket_5_qty': 0.0, 'bracket_5_val': 0.0,
                        'total_qty': 0.0, 'total_val': 0.0,
                    }
                }
            grouped[categ]['lines'].append(item)
            for k in grand_total.keys():
                if not k.endswith('_price'):
                    grouped[categ]['total'][k] += item.get(k, 0.0)
                    grand_total[k] += item.get(k, 0.0)
            grouped[categ]['total']['unit_cost'] += item.get('unit_cost', 0.0)

        for categ_name, cat_data in grouped.items():
            tot = cat_data['total']
            for b in range(1, 6):
                qty_k = f"bracket_{b}_qty"
                val_k = f"bracket_{b}_val"
                price_k = f"bracket_{b}_price"
                tot[price_k] = round(tot[val_k] / tot[qty_k], 2) if tot[qty_k] > 0 else 0.0

        for b in range(1, 6):
            qty_k = f"bracket_{b}_qty"
            val_k = f"bracket_{b}_val"
            price_k = f"bracket_{b}_price"
            grand_total[price_k] = round(grand_total[val_k] / grand_total[qty_k], 2) if grand_total[qty_k] > 0 else 0.0

        return {
            'grouped': grouped,
            'grand_total': grand_total,
            'brackets': self._get_bracket_labels(),
        }

    def _get_or_create_period_views(self):
        """Finds or creates dedicated List and Pivot ir.ui.view records for the current period_length."""
        self.ensure_one()
        p = self.period_length
        brackets = self._get_bracket_labels()
        b1, b2, b3, b4, b5 = brackets['b1'], brackets['b2'], brackets['b3'], brackets['b4'], brackets['b5']

        View = self.env['ir.ui.view'].sudo()

        # 1. Dynamic List View per period_length
        list_view_name = f"inventory.aging.report.line.list.p{p}"
        list_view = View.search([
            ('model', '=', 'inventory.aging.report.line'),
            ('type', '=', 'list'),
            ('name', '=', list_view_name),
        ], limit=1)

        list_arch = f"""<list string="Inventory Aging and Valuation Report" create="false" edit="false" delete="false">
            <header>
                <button name="action_refresh_report" string="Refresh Data" type="object" class="btn-primary" icon="fa-refresh"/>
                <button name="action_print_pdf" string="Print PDF" type="object" class="btn-secondary" icon="fa-file-pdf-o"/>
                <button name="action_export_excel" string="Export Excel" type="object" class="btn-secondary" icon="fa-file-excel-o"/>
            </header>
            <field name="default_code" string="Internal Reference"/>
            <field name="product_id"/>
            <field name="categ_id"/>
            <field name="warehouse_id"/>
            <field name="unit_cost" widget="monetary" options="{{'currency_field': 'company_id'}}"/>
            <field name="total_qty" string="Total On-Hand Qty" sum="Grand Total Qty" class="fw-bold"/>
            <field name="total_val" string="Total Value" sum="Grand Total Value" widget="monetary" class="fw-bold"/>
            <field name="bracket_1_price" string="{b1} Price" widget="monetary" optional="show"/>
            <field name="bracket_1_qty" string="{b1} Qty" sum="Total"/>
            <field name="bracket_1_val" string="{b1} Value" sum="Total" widget="monetary"/>
            <field name="bracket_2_price" string="{b2} Price" widget="monetary" optional="show"/>
            <field name="bracket_2_qty" string="{b2} Qty" sum="Total"/>
            <field name="bracket_2_val" string="{b2} Value" sum="Total" widget="monetary"/>
            <field name="bracket_3_price" string="{b3} Price" widget="monetary" optional="show"/>
            <field name="bracket_3_qty" string="{b3} Qty" sum="Total"/>
            <field name="bracket_3_val" string="{b3} Value" sum="Total" widget="monetary"/>
            <field name="bracket_4_price" string="{b4} Price" widget="monetary" optional="show"/>
            <field name="bracket_4_qty" string="{b4} Qty" sum="Total"/>
            <field name="bracket_4_val" string="{b4} Value" sum="Total" widget="monetary"/>
            <field name="bracket_5_price" string="{b5} Price" widget="monetary" optional="show"/>
            <field name="bracket_5_qty" string="{b5} Qty" sum="Total"/>
            <field name="bracket_5_val" string="{b5} Value" sum="Total" widget="monetary"/>
            <field name="company_id" column_invisible="True"/>
        </list>"""

        if list_view:
            list_view.write({'arch': list_arch})
        else:
            list_view = View.create({
                'name': list_view_name,
                'model': 'inventory.aging.report.line',
                'type': 'list',
                'arch': list_arch,
                'priority': 1,
            })

        # 2. Dynamic Pivot View per period_length
        pivot_view_name = f"inventory.aging.report.line.pivot.p{p}"
        pivot_view = View.search([
            ('model', '=', 'inventory.aging.report.line'),
            ('type', '=', 'pivot'),
            ('name', '=', pivot_view_name),
        ], limit=1)

        pivot_arch = f"""<pivot string="Inventory Aging and Valuation Report" display_quantity="false">
            <field name="company_id" type="row"/>
            <field name="categ_id" type="row"/>
            <field name="product_id" type="row"/>
            <field name="bracket_1_price" string="{b1} Price" type="measure"/>
            <field name="bracket_1_qty" string="{b1} Qty" type="measure"/>
            <field name="bracket_1_val" string="{b1} Value" type="measure"/>
            <field name="bracket_2_price" string="{b2} Price" type="measure"/>
            <field name="bracket_2_qty" string="{b2} Qty" type="measure"/>
            <field name="bracket_2_val" string="{b2} Value" type="measure"/>
            <field name="bracket_3_price" string="{b3} Price" type="measure"/>
            <field name="bracket_3_qty" string="{b3} Qty" type="measure"/>
            <field name="bracket_3_val" string="{b3} Value" type="measure"/>
            <field name="bracket_4_price" string="{b4} Price" type="measure"/>
            <field name="bracket_4_qty" string="{b4} Qty" type="measure"/>
            <field name="bracket_4_val" string="{b4} Value" type="measure"/>
            <field name="bracket_5_price" string="{b5} Price" type="measure"/>
            <field name="bracket_5_qty" string="{b5} Qty" type="measure"/>
            <field name="bracket_5_val" string="{b5} Value" type="measure"/>
            <field name="total_qty" string="Total Qty" type="measure"/>
            <field name="total_val" string="Total Value" type="measure"/>
        </pivot>"""

        if pivot_view:
            pivot_view.write({'arch': pivot_arch})
        else:
            pivot_view = View.create({
                'name': pivot_view_name,
                'model': 'inventory.aging.report.line',
                'type': 'pivot',
                'arch': pivot_arch,
                'priority': 1,
            })

        self.env.invalidate_all()
        return list_view, pivot_view

    def _recalculate_report_lines(self):
        """Recalculate live stock data and synchronize report lines in-place without deleting existing record IDs."""
        Line = self.env['inventory.aging.report.line'].sudo()
        for wiz in self:
            data_lines = wiz._get_aging_data()
            existing_lines = Line.search([('wizard_id', '=', wiz.id)])
            existing_map = {
                (l.product_id.id, l.warehouse_id.id if l.warehouse_id else False): l
                for l in existing_lines
            }
            seen_keys = set()
            to_create = []

            for line in data_lines:
                key = (line['product_id'], line.get('warehouse_id') or False)
                seen_keys.add(key)
                vals = {
                    'wizard_id': wiz.id,
                    'product_id': line['product_id'],
                    'default_code': line['default_code'],
                    'categ_id': line['categ_id'],
                    'warehouse_id': line['warehouse_id'],
                    'company_id': wiz.company_id.id,
                    'unit_cost': line['unit_cost'],
                    'bracket_1_price': line['bracket_1_price'],
                    'bracket_1_qty': line['bracket_1_qty'],
                    'bracket_1_val': line['bracket_1_val'],
                    'bracket_2_price': line['bracket_2_price'],
                    'bracket_2_qty': line['bracket_2_qty'],
                    'bracket_2_val': line['bracket_2_val'],
                    'bracket_3_price': line['bracket_3_price'],
                    'bracket_3_qty': line['bracket_3_qty'],
                    'bracket_3_val': line['bracket_3_val'],
                    'bracket_4_price': line['bracket_4_price'],
                    'bracket_4_qty': line['bracket_4_qty'],
                    'bracket_4_val': line['bracket_4_val'],
                    'bracket_5_price': line['bracket_5_price'],
                    'bracket_5_qty': line['bracket_5_qty'],
                    'bracket_5_val': line['bracket_5_val'],
                    'total_qty': line['total_qty'],
                    'total_val': line['total_val'],
                }

                if key in existing_map:
                    existing_line = existing_map[key]
                    write_vals = {
                        k: v for k, v in vals.items()
                        if k not in ('wizard_id', 'product_id', 'warehouse_id', 'company_id')
                        and existing_line[k] != v
                    }
                    if write_vals:
                        existing_line.write(write_vals)
                else:
                    to_create.append(vals)

            if to_create:
                Line.create(to_create)

            stale_lines = [l for k, l in existing_map.items() if k not in seen_keys]
            if stale_lines:
                Line.browse([l.id for l in stale_lines]).unlink()

    # -------------------------------------------------------------------------
    # 1. VIEW REPORT ON SCREEN
    # -------------------------------------------------------------------------
    def action_view_report(self):
        """Calculates data and opens the report lines in List, Pivot, and Graph views."""
        self.ensure_one()

        self._recalculate_report_lines()

        # Get or create period-specific views so Odoo JS view cache never collides across period lengths
        list_view, pivot_view = self._get_or_create_period_views()

        graph_view = self.env.ref('micra_inventory_ageing_report.view_inventory_aging_report_line_graph', raise_if_not_found=False)
        search_view = self.env.ref('micra_inventory_ageing_report.view_inventory_aging_report_line_search', raise_if_not_found=False)

        action = {
            'name': _("Inventory Aging and Valuation Report (%s Days)") % self.period_length,
            'type': 'ir.actions.act_window',
            'res_model': 'inventory.aging.report.line',
            'view_mode': 'list,pivot,graph',
            'views': [
                (list_view.id, 'list'),
                (pivot_view.id, 'pivot'),
                (graph_view.id if graph_view else False, 'graph'),
            ],
            'search_view_id': [search_view.id, search_view.name] if search_view else False,
            'domain': [('wizard_id', '=', self.id)],
            'context': {
                'search_default_group_by_categ': 1,
                'search_default_has_stock': 1,
                'period_length': self.period_length,
                'default_wizard_id': self.id,
            }
        }
        return action

    # -------------------------------------------------------------------------
    # 2. PRINT PDF REPORT
    # -------------------------------------------------------------------------
    def action_print_pdf(self):
        """Generates and downloads the QWeb PDF report."""
        self.ensure_one()
        return self.env.ref('micra_inventory_ageing_report.action_report_inventory_aging').report_action(self)

    # -------------------------------------------------------------------------
    # 3. EXPORT EXCEL REPORT
    # -------------------------------------------------------------------------
    def action_export_excel(self):
        """Generates and downloads an Excel (.xlsx) file with full report details."""
        self.ensure_one()
        grouped_data = self._get_aging_data_grouped()
        brackets = grouped_data['brackets']

        output = io.BytesIO()

        if xlsxwriter:
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            sheet = workbook.add_worksheet('Inventory Aging and Valuation Report')

            # Formats
            title_format = workbook.add_format({
                'bold': True, 'font_size': 16, 'font_color': '#4A154B',
                'align': 'left', 'valign': 'vcenter'
            })
            filter_hdr_format = workbook.add_format({
                'bold': True, 'font_size': 10, 'bg_color': '#F2F2F2',
                'font_color': '#333333', 'border': 1
            })
            filter_val_format = workbook.add_format({
                'font_size': 10, 'border': 1
            })
            table_header_format = workbook.add_format({
                'bold': True, 'font_size': 10, 'bg_color': '#69558C',
                'font_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter',
                'border': 1, 'text_wrap': True
            })
            categ_header_format = workbook.add_format({
                'bold': True, 'font_size': 11, 'bg_color': '#EFEBF5',
                'font_color': '#2C1B4D', 'border': 1
            })
            data_str_format = workbook.add_format({
                'font_size': 10, 'border': 1, 'valign': 'vcenter'
            })
            data_num_format = workbook.add_format({
                'font_size': 10, 'border': 1, 'num_format': '#,##0.00', 'valign': 'vcenter'
            })
            categ_total_format = workbook.add_format({
                'bold': True, 'font_size': 10, 'bg_color': '#F7F5FA',
                'border': 1, 'num_format': '#,##0.00', 'valign': 'vcenter'
            })
            grand_total_format = workbook.add_format({
                'bold': True, 'font_size': 11, 'bg_color': '#69558C',
                'font_color': '#FFFFFF', 'border': 1, 'num_format': '#,##0.00', 'valign': 'vcenter'
            })

            # Title
            sheet.write(0, 0, 'Inventory Aging and Valuation Report', title_format)
            sheet.set_row(0, 30)

            # Filter Summary
            sheet.write(2, 0, 'Aging Start Date', filter_hdr_format)
            sheet.write(2, 1, str(self.aging_date), filter_val_format)
            sheet.write(2, 3, 'Period Length', filter_hdr_format)
            sheet.write(2, 4, f"{self.period_length} Days", filter_val_format)

            sheet.write(3, 0, 'Company', filter_hdr_format)
            sheet.write(3, 1, self.company_id.name, filter_val_format)
            sheet.write(3, 3, 'Filter By', filter_hdr_format)
            sheet.write(3, 4, dict(self._fields['filter_by'].selection).get(self.filter_by, ''), filter_val_format)

            # Table Headers
            headers = [
                'Internal Reference', 'Product', 'Category', 'Warehouse', 'Unit Cost',
                'Total On-Hand Qty', 'Total Value',
                f"{brackets['b1']} Price", f"{brackets['b1']} Qty", f"{brackets['b1']} Value",
                f"{brackets['b2']} Price", f"{brackets['b2']} Qty", f"{brackets['b2']} Value",
                f"{brackets['b3']} Price", f"{brackets['b3']} Qty", f"{brackets['b3']} Value",
                f"{brackets['b4']} Price", f"{brackets['b4']} Qty", f"{brackets['b4']} Value",
                f"{brackets['b5']} Price", f"{brackets['b5']} Qty", f"{brackets['b5']} Value",
            ]

            row = 5
            sheet.set_row(row, 28)
            for col, h in enumerate(headers):
                sheet.write(row, col, h, table_header_format)

            row += 1
            for categ_name, cat_data in grouped_data['grouped'].items():
                # Category Header row
                sheet.write(row, 0, f"▸ {categ_name} ({len(cat_data['lines'])})", categ_header_format)
                for c in range(1, len(headers)):
                    sheet.write(row, c, '', categ_header_format)
                row += 1

                for line in cat_data['lines']:
                    sheet.write(row, 0, line['default_code'], data_str_format)
                    sheet.write(row, 1, line['product_name'], data_str_format)
                    sheet.write(row, 2, line['categ_name'], data_str_format)
                    sheet.write(row, 3, line['warehouse_name'], data_str_format)
                    sheet.write(row, 4, line['unit_cost'], data_num_format)
                    sheet.write(row, 5, line['total_qty'], data_num_format)
                    sheet.write(row, 6, line['total_val'], data_num_format)
                    sheet.write(row, 7, line['bracket_1_price'], data_num_format)
                    sheet.write(row, 8, line['bracket_1_qty'], data_num_format)
                    sheet.write(row, 9, line['bracket_1_val'], data_num_format)
                    sheet.write(row, 10, line['bracket_2_price'], data_num_format)
                    sheet.write(row, 11, line['bracket_2_qty'], data_num_format)
                    sheet.write(row, 12, line['bracket_2_val'], data_num_format)
                    sheet.write(row, 13, line['bracket_3_price'], data_num_format)
                    sheet.write(row, 14, line['bracket_3_qty'], data_num_format)
                    sheet.write(row, 15, line['bracket_3_val'], data_num_format)
                    sheet.write(row, 16, line['bracket_4_price'], data_num_format)
                    sheet.write(row, 17, line['bracket_4_qty'], data_num_format)
                    sheet.write(row, 18, line['bracket_4_val'], data_num_format)
                    sheet.write(row, 19, line['bracket_5_price'], data_num_format)
                    sheet.write(row, 20, line['bracket_5_qty'], data_num_format)
                    sheet.write(row, 21, line['bracket_5_val'], data_num_format)
                    row += 1

                # Category Subtotal row
                tot = cat_data['total']
                sheet.write(row, 0, f"Subtotal - {categ_name}", categ_total_format)
                sheet.write(row, 1, '', categ_total_format)
                sheet.write(row, 2, '', categ_total_format)
                sheet.write(row, 3, '', categ_total_format)
                sheet.write(row, 4, tot['unit_cost'], categ_total_format)
                sheet.write(row, 5, tot['total_qty'], categ_total_format)
                sheet.write(row, 6, tot['total_val'], categ_total_format)
                sheet.write(row, 7, tot['bracket_1_price'], categ_total_format)
                sheet.write(row, 8, tot['bracket_1_qty'], categ_total_format)
                sheet.write(row, 9, tot['bracket_1_val'], categ_total_format)
                sheet.write(row, 10, tot['bracket_2_price'], categ_total_format)
                sheet.write(row, 11, tot['bracket_2_qty'], categ_total_format)
                sheet.write(row, 12, tot['bracket_2_val'], categ_total_format)
                sheet.write(row, 13, tot['bracket_3_price'], categ_total_format)
                sheet.write(row, 14, tot['bracket_3_qty'], categ_total_format)
                sheet.write(row, 15, tot['bracket_3_val'], categ_total_format)
                sheet.write(row, 16, tot['bracket_4_price'], categ_total_format)
                sheet.write(row, 17, tot['bracket_4_qty'], categ_total_format)
                sheet.write(row, 18, tot['bracket_4_val'], categ_total_format)
                sheet.write(row, 19, tot['bracket_5_price'], categ_total_format)
                sheet.write(row, 20, tot['bracket_5_qty'], categ_total_format)
                sheet.write(row, 21, tot['bracket_5_val'], categ_total_format)
                row += 1

            # Grand Total row
            gt = grouped_data['grand_total']
            sheet.write(row, 0, 'Grand Total', grand_total_format)
            sheet.write(row, 1, '', grand_total_format)
            sheet.write(row, 2, '', grand_total_format)
            sheet.write(row, 3, '', grand_total_format)
            sheet.write(row, 4, '', grand_total_format)
            sheet.write(row, 5, gt['total_qty'], grand_total_format)
            sheet.write(row, 6, gt['total_val'], grand_total_format)
            sheet.write(row, 7, gt['bracket_1_price'], grand_total_format)
            sheet.write(row, 8, gt['bracket_1_qty'], grand_total_format)
            sheet.write(row, 9, gt['bracket_1_val'], grand_total_format)
            sheet.write(row, 10, gt['bracket_2_price'], grand_total_format)
            sheet.write(row, 11, gt['bracket_2_qty'], grand_total_format)
            sheet.write(row, 12, gt['bracket_2_val'], grand_total_format)
            sheet.write(row, 13, gt['bracket_3_price'], grand_total_format)
            sheet.write(row, 14, gt['bracket_3_qty'], grand_total_format)
            sheet.write(row, 15, gt['bracket_3_val'], grand_total_format)
            sheet.write(row, 16, gt['bracket_4_price'], grand_total_format)
            sheet.write(row, 17, gt['bracket_4_qty'], grand_total_format)
            sheet.write(row, 18, gt['bracket_4_val'], grand_total_format)
            sheet.write(row, 19, gt['bracket_5_price'], grand_total_format)
            sheet.write(row, 20, gt['bracket_5_qty'], grand_total_format)
            sheet.write(row, 21, gt['bracket_5_val'], grand_total_format)

            # Auto fit column widths
            sheet.set_column(0, 0, 18)
            sheet.set_column(1, 1, 32)
            sheet.set_column(2, 2, 22)
            sheet.set_column(3, 3, 16)
            sheet.set_column(4, 4, 13)
            sheet.set_column(5, 5, 18)
            sheet.set_column(6, 6, 16)
            sheet.set_column(7, 21, 14)

            workbook.close()
        else:
            # Fallback CSV/Spreadsheet generation if xlsxwriter is not in environment
            import csv
            str_out = io.StringIO()
            writer = csv.writer(str_out)
            writer.writerow(['Inventory Aging and Valuation Report'])
            writer.writerow(['Aging Start Date', str(self.aging_date), 'Period Length', f"{self.period_length} Days"])
            writer.writerow(['Company', self.company_id.name, 'Filter By', self.filter_by])
            writer.writerow([])
            headers = [
                'Internal Reference', 'Product', 'Category', 'Warehouse', 'Unit Cost',
                'Total On-Hand Qty', 'Total Value',
                f"{brackets['b1']} Price", f"{brackets['b1']} Qty", f"{brackets['b1']} Value",
                f"{brackets['b2']} Price", f"{brackets['b2']} Qty", f"{brackets['b2']} Value",
                f"{brackets['b3']} Price", f"{brackets['b3']} Qty", f"{brackets['b3']} Value",
                f"{brackets['b4']} Price", f"{brackets['b4']} Qty", f"{brackets['b4']} Value",
                f"{brackets['b5']} Price", f"{brackets['b5']} Qty", f"{brackets['b5']} Value",
            ]
            writer.writerow(headers)
            for categ_name, cat_data in grouped_data['grouped'].items():
                for line in cat_data['lines']:
                    writer.writerow([
                        line['default_code'], line['product_name'], line['categ_name'], line['warehouse_name'],
                        line['unit_cost'], line['total_qty'], line['total_val'],
                        line['bracket_1_price'], line['bracket_1_qty'], line['bracket_1_val'],
                        line['bracket_2_price'], line['bracket_2_qty'], line['bracket_2_val'],
                        line['bracket_3_price'], line['bracket_3_qty'], line['bracket_3_val'],
                        line['bracket_4_price'], line['bracket_4_qty'], line['bracket_4_val'],
                        line['bracket_5_price'], line['bracket_5_qty'], line['bracket_5_val'],
                    ])
            gt = grouped_data['grand_total']
            writer.writerow([
                'Grand Total', '', '', '', '',
                gt['total_qty'], gt['total_val'],
                gt['bracket_1_price'], gt['bracket_1_qty'], gt['bracket_1_val'],
                gt['bracket_2_price'], gt['bracket_2_qty'], gt['bracket_2_val'],
                gt['bracket_3_price'], gt['bracket_3_qty'], gt['bracket_3_val'],
                gt['bracket_4_price'], gt['bracket_4_qty'], gt['bracket_4_val'],
                gt['bracket_5_price'], gt['bracket_5_qty'], gt['bracket_5_val'],
            ])
            output.write(str_out.getvalue().encode('utf-8'))

        output.seek(0)
        file_data = output.read()

        attachment = self.env['ir.attachment'].create({
            'name': f"Inventory_Aging_Report_{self.aging_date}.xlsx",
            'type': 'binary',
            'raw': file_data,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'res_model': self._name,
            'res_id': self.id,
        })

        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'self',
        }


class InventoryAgingReportLine(models.TransientModel):
    _name = 'inventory.aging.report.line'
    _description = 'Inventory Aging Report Line'
    _order = 'categ_id, product_id'

    wizard_id = fields.Many2one(
        'inventory.aging.report.wizard',
        string="Wizard",
        ondelete='cascade',
        index=True
    )
    product_id = fields.Many2one(
        'product.product',
        string="Product",
        readonly=True
    )
    default_code = fields.Char(
        string="Internal Reference",
        readonly=True
    )
    categ_id = fields.Many2one(
        'product.category',
        string="Category",
        readonly=True
    )
    warehouse_id = fields.Many2one(
        'stock.warehouse',
        string="Warehouse",
        readonly=True
    )
    company_id = fields.Many2one(
        'res.company',
        string="Company",
        readonly=True
    )
    unit_cost = fields.Float(
        string="Unit Cost",
        readonly=True,
        digits='Product Price'
    )
    bracket_1_price = fields.Float(string="Bracket 1 Price", readonly=True, digits='Product Price')
    bracket_1_qty = fields.Float(string="Bracket 1 Qty", readonly=True)
    bracket_1_val = fields.Float(string="Bracket 1 Value", readonly=True)
    bracket_2_price = fields.Float(string="Bracket 2 Price", readonly=True, digits='Product Price')
    bracket_2_qty = fields.Float(string="Bracket 2 Qty", readonly=True)
    bracket_2_val = fields.Float(string="Bracket 2 Value", readonly=True)
    bracket_3_price = fields.Float(string="Bracket 3 Price", readonly=True, digits='Product Price')
    bracket_3_qty = fields.Float(string="Bracket 3 Qty", readonly=True)
    bracket_3_val = fields.Float(string="Bracket 3 Value", readonly=True)
    bracket_4_price = fields.Float(string="Bracket 4 Price", readonly=True, digits='Product Price')
    bracket_4_qty = fields.Float(string="Bracket 4 Qty", readonly=True)
    bracket_4_val = fields.Float(string="Bracket 4 Value", readonly=True)
    bracket_5_price = fields.Float(string="Bracket 5 Price", readonly=True, digits='Product Price')
    bracket_5_qty = fields.Float(string="Bracket 5 Qty", readonly=True)
    bracket_5_val = fields.Float(string="Bracket 5 Value", readonly=True)
    total_qty = fields.Float(string="Total Qty", readonly=True)
    total_val = fields.Float(string="Total Value", readonly=True)

    @api.model
    def _get_period_length_from_context(self):
        p = self.env.context.get('period_length')
        if not p and self.env.context.get('default_wizard_id'):
            wiz = self.env['inventory.aging.report.wizard'].browse(self.env.context.get('default_wizard_id'))
            if wiz.exists():
                p = wiz.period_length
        if not p and self.env.context.get('active_id') and self.env.context.get('active_model') == 'inventory.aging.report.wizard':
            wiz = self.env['inventory.aging.report.wizard'].browse(self.env.context.get('active_id'))
            if wiz.exists():
                p = wiz.period_length
        if not p and self.env.context.get('domain'):
            domain = self.env.context.get('domain')
            if isinstance(domain, list):
                for item in domain:
                    if isinstance(item, (list, tuple)) and len(item) == 3 and item[0] == 'wizard_id' and item[1] == '=':
                        wiz = self.env['inventory.aging.report.wizard'].browse(item[2])
                        if wiz.exists():
                            p = wiz.period_length
                            break
        if not p:
            latest_wiz = self.env['inventory.aging.report.wizard'].search([], order='id desc', limit=1)
            if latest_wiz:
                p = latest_wiz.period_length
        return p or 30

    @api.model
    def fields_get(self, allfields=None, attributes=None):
        res = super().fields_get(allfields=allfields, attributes=attributes)
        p = self._get_period_length_from_context()
        brackets = self.env['inventory.aging.report.wizard']._compute_bracket_labels_dict(p)
        b1, b2, b3, b4, b5 = brackets['b1'], brackets['b2'], brackets['b3'], brackets['b4'], brackets['b5']

        labels = {
            'bracket_1_price': f"{b1} Price",
            'bracket_1_qty': f"{b1} Qty",
            'bracket_1_val': f"{b1} Value",
            'bracket_2_price': f"{b2} Price",
            'bracket_2_qty': f"{b2} Qty",
            'bracket_2_val': f"{b2} Value",
            'bracket_3_price': f"{b3} Price",
            'bracket_3_qty': f"{b3} Qty",
            'bracket_3_val': f"{b3} Value",
            'bracket_4_price': f"{b4} Price",
            'bracket_4_qty': f"{b4} Qty",
            'bracket_4_val': f"{b4} Value",
            'bracket_5_price': f"{b5} Price",
            'bracket_5_qty': f"{b5} Qty",
            'bracket_5_val': f"{b5} Value",
        }
        for fname, label in labels.items():
            if fname in res:
                res[fname]['string'] = label
        return res

    @api.model
    def get_views(self, views, options=None):
        res = super().get_views(views, options=options)
        p = self._get_period_length_from_context()
        brackets = self.env['inventory.aging.report.wizard']._compute_bracket_labels_dict(p)
        b1, b2, b3, b4, b5 = brackets['b1'], brackets['b2'], brackets['b3'], brackets['b4'], brackets['b5']

        labels = {
            'bracket_1_price': f"{b1} Price",
            'bracket_1_qty': f"{b1} Qty",
            'bracket_1_val': f"{b1} Value",
            'bracket_2_price': f"{b2} Price",
            'bracket_2_qty': f"{b2} Qty",
            'bracket_2_val': f"{b2} Value",
            'bracket_3_price': f"{b3} Price",
            'bracket_3_qty': f"{b3} Qty",
            'bracket_3_val': f"{b3} Value",
            'bracket_4_price': f"{b4} Price",
            'bracket_4_qty': f"{b4} Qty",
            'bracket_4_val': f"{b4} Value",
            'bracket_5_price': f"{b5} Price",
            'bracket_5_qty': f"{b5} Qty",
            'bracket_5_val': f"{b5} Value",
        }

        # Dynamically modify arch XML list so OWL list renderer uses new field headers
        views_dict = res.get('views', {})
        for vname, vdata in views_dict.items():
            if isinstance(vdata, dict):
                if 'arch' in vdata:
                    try:
                        from lxml import etree
                        root = etree.fromstring(vdata['arch'].encode('utf-8'))
                        updated = False
                        for node in root.xpath('//field'):
                            fname = node.get('name')
                            if fname in labels:
                                node.set('string', labels[fname])
                                updated = True
                        if updated:
                            vdata['arch'] = etree.tostring(root, encoding='unicode')
                    except Exception as e:
                        _logger.warning("Failed to update dynamic arch labels: %s", e)

                if 'fields' in vdata:
                    for fname, label in labels.items():
                        if fname in vdata['fields']:
                            vdata['fields'][fname]['string'] = label

        models_dict = res.get('models', {})
        if self._name in models_dict:
            target_fields = models_dict[self._name]
            for fname, label in labels.items():
                if fname in target_fields:
                    target_fields[fname]['string'] = label

        return res

    def _get_wizards(self):
        """Helper to find the associated wizard record(s) safely in any context."""
        valid_lines = self.exists() if self else self.env['inventory.aging.report.line']
        wizards = valid_lines.mapped('wizard_id').exists()

        if not wizards:
            ctx = self.env.context
            wid = ctx.get('default_wizard_id')
            if not wid and ctx.get('active_model') == 'inventory.aging.report.wizard':
                wid = ctx.get('active_id') or (ctx.get('active_ids')[0] if ctx.get('active_ids') else None)

            if not wid:
                domain = ctx.get('domain') or []
                if isinstance(domain, list):
                    for item in domain:
                        if isinstance(item, (list, tuple)) and len(item) == 3 and item[0] == 'wizard_id' and item[1] == '=':
                            wid = item[2]
                            break

            if not wid and ctx.get('params', {}).get('default_wizard_id'):
                wid = ctx.get('params', {}).get('default_wizard_id')

            if wid:
                wizards = self.env['inventory.aging.report.wizard'].browse(wid).exists()

            if not wizards:
                wizards = self.env['inventory.aging.report.wizard'].search([], order='id desc', limit=1)

        return wizards

    def action_refresh_report(self):
        """Action button on the report view to manually trigger instant recalculation."""
        wizards = self._get_wizards()
        for wiz in wizards:
            wiz._recalculate_report_lines()
        return {
            'type': 'ir.actions.client',
            'tag': 'reload',
        }

    def action_print_pdf(self):
        """Triggers QWeb PDF report printing from report line view or Action gear menu."""
        wizards = self._get_wizards()
        if not wizards:
            raise UserError(_("No active Inventory Aging and Valuation Report wizard found to print PDF."))
        return wizards[0].action_print_pdf()

    def action_export_excel(self):
        """Triggers Excel report export from report line view or Action gear menu."""
        wizards = self._get_wizards()
        if not wizards:
            raise UserError(_("No active Inventory Aging and Valuation Report wizard found to export Excel."))
        return wizards[0].action_export_excel()

    @api.model
    def _refresh_wizard_lines(self, domain=None):
        """Find wizard_id and recalculate live report data before view retrieval."""
        if self.env.context.get('skip_aging_refresh'):
            return
        wizard_id = self.env.context.get('default_wizard_id')
        if not wizard_id and self.env.context.get('active_model') == 'inventory.aging.report.wizard':
            wizard_id = self.env.context.get('active_id')
        if not wizard_id and domain:
            for item in domain:
                if isinstance(item, (list, tuple)) and len(item) == 3 and item[0] == 'wizard_id' and item[1] == '=':
                    wizard_id = item[2]
                    break
        if not wizard_id and self.env.context.get('params', {}).get('default_wizard_id'):
            wizard_id = self.env.context.get('params', {}).get('default_wizard_id')

        if wizard_id:
            wiz = self.env['inventory.aging.report.wizard'].browse(wizard_id)
            if wiz.exists():
                wiz.with_context(skip_aging_refresh=True)._recalculate_report_lines()

    @api.model
    def search_read(self, domain=None, fields=None, offset=0, limit=None, order=None, **kwargs):
        self._refresh_wizard_lines(domain=domain)
        return super().search_read(domain=domain, fields=fields, offset=offset, limit=limit, order=order, **kwargs)

    @api.model
    def read_group(self, domain, fields, groupby, offset=0, limit=None, orderby=False, lazy=True):
        self._refresh_wizard_lines(domain=domain)
        return super().read_group(domain, fields, groupby, offset=offset, limit=limit, orderby=orderby, lazy=lazy)

    @api.model
    def _read_group(self, domain, groupby=(), aggregates=(), having=(), offset=0, limit=None, order=None):
        self._refresh_wizard_lines(domain=domain)
        return super()._read_group(domain, groupby=groupby, aggregates=aggregates, having=having, offset=offset, limit=limit, order=order)