# -*- coding: utf-8 -*-

from odoo import models, api, _
from odoo.exceptions import UserError


class ReportInventoryAging(models.AbstractModel):
    _name = 'report.micra_inventory_ageing_report.report_aging_pdf'
    _description = 'Inventory Aging Report QWeb'

    @api.model
    def _get_report_values(self, docids, data=None):
        Wizard = self.env['inventory.aging.report.wizard']
        Line = self.env['inventory.aging.report.line']

        # 1. Try to resolve docids as Wizard records
        docs = Wizard.browse(docids).exists()

        # 2. If docids are line IDs, resolve the parent wizard(s)
        if not docs:
            lines = Line.browse(docids).exists()
            if lines:
                docs = lines.mapped('wizard_id').exists()

        # 3. Fallback to context default_wizard_id or active_id
        if not docs:
            wid = self.env.context.get('default_wizard_id')
            if not wid and self.env.context.get('active_model') == 'inventory.aging.report.wizard':
                wid = self.env.context.get('active_id')
            if wid:
                docs = Wizard.browse(wid).exists()

        # 4. Fallback to the latest created wizard
        if not docs:
            docs = Wizard.search([], order='id desc', limit=1)

        if not docs:
            raise UserError(_("No active Inventory Aging and Valuation Report wizard found to print PDF."))

        return {
            'doc_ids': docs.ids,
            'doc_model': 'inventory.aging.report.wizard',
            'docs': docs,
        }
