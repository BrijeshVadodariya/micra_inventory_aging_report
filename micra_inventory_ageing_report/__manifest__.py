{
    'name': "Incoming Stock Valuation Report",
    
    'summary': "Incoming Stock Valuation & FIFO Receipt Analysis with dynamic period brackets, purchase price tracking, multi-warehouse filtering, interactive List/Pivot/Graph views, PDF & formatted Excel (.xlsx) export.",

    'description':"""
        Incoming Stock Valuation Report for Odoo
        ========================================
        Track incoming stock valuation, analyze purchase receipts across dynamic periods, and evaluate accurate inventory valuation with true FIFO cost allocation.

        Key Features & Highlights
        -------------------------
        * Accurate FIFO Inbound Valuation:
        - Evaluates on-hand quantities against incoming stock moves (newest to oldest).
        - Computes weighted average unit prices and total values across period intervals.

        * Dynamic & Configurable Period Brackets:
        - Fully customizable period length (e.g., 30, 60, 90 days or daily breakdown).
        - Automatically updates table headers and bracket labels dynamically (e.g., 0-30, 31-60, 61-90, 91-120, 121+ Days).

        * Purchase Price Tracking:
        - Direct visibility of previous purchase prices and price growth percentage on purchase orders.

        * Multi-Dimensional Filters:
        - Filter stock by Company, Warehouse(s), Internal Stock Location(s) (including child locations), Specific Products, or Product Categories.

        * Interactive On-Screen Analytics:
        - Dynamic List View: Column headers update per period with live subtotals, grand totals, and quick action buttons.
        - Pivot & Matrix Analysis: Multi-level breakdowns by Company, Category, and Product with quantity/valuation measures.
        - Graphical Distribution: Visualize incoming valuation trends directly inside Odoo.

        * Executive Reports & Export:
        - Formatted Excel (.xlsx) Export: Branded spreadsheet with category groupings, subtotals, and total stock valuations.
        - Printable QWeb PDF Report: Clean, presentation-ready PDF report organized by product category.
    """,

    'author': "Micra Digital",
    'website': "www.micra.digital",
    'category': 'Inventory/Inventory',
    'version': '19.0.1.0.6',
    'installable': True,
    'application': True,
    'license': 'OPL-1',
    'price': 11.00,
    'currency': 'USD',

    # any module necessary for this one to work correctly
    'depends': ['base', 'stock', 'purchase'],

    # always loaded
    'data': [
        'security/ir.model.access.csv',
        'views/purchase_order_views.xml',
        'wizard/inventory_aging_report_wizard_views.xml',
        'report/inventory_aging_report_template.xml',
    ],

    'images': ['static/description/banner.png'],

}

