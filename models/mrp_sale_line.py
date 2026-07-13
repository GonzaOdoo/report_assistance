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