from odoo import api, fields, models


class MrpProduction(models.Model):
    _inherit = "mrp.production"

    source_sale_line_id = fields.Many2one(
        "sale.order.line",
        string="Source Sale Line",
        compute="_compute_source_sale_line_id",
        store=True,
    )
    line_description = fields.Text(string="Descripción linea de venta",related="source_sale_line_id.name",store=True)

    sale_order_names = fields.Char(
        string="Venta(Origen)",
        compute="_compute_sale_order_names",
        store=True,
    )

    @api.depends(
        "reference_ids.sale_ids.name",
        "sale_line_id.order_id.name",
    )
    def _compute_sale_order_names(self):
        for production in self:
            sale_orders = production.reference_ids.sale_ids | production.sale_line_id.order_id
            # Evita duplicados y mantiene un orden consistente
            names = sorted(set(sale_orders.mapped("name")))
            production.sale_order_names = ", ".join(names)

    @api.depends(
        "sale_line_id",
        "production_group_id.parent_ids.production_ids.sale_line_id",
    )
    def _compute_source_sale_line_id(self):
        for production in self:
            production.source_sale_line_id = (
                production.sale_line_id
                or production._get_sources()[:1].sale_line_id
            )



class StockPicking(models.Model):
    _inherit = "stock.picking"

    sale_order_names = fields.Char(
        string="Venta(Origen)",
        compute="_compute_sale_order_names",
        store=True,
    )

    @api.depends("origin")
    def _compute_sale_order_names(self):
        MrpProduction = self.env["mrp.production"]

        for picking in self:
            picking.sale_order_names = picking.origin or ""

            if not picking.origin:
                continue

            production = MrpProduction.search([("name", "=", picking.origin)], limit=1)
            if not production:
                continue

            sale_orders = production.reference_ids.sale_ids | production.sale_line_id.order_id
            if sale_orders:
                picking.sale_order_names = ", ".join(
                    sorted(set(sale_orders.mapped("name")))
                )