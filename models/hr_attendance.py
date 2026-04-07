from odoo import models, fields, api
from datetime import date,datetime,timedelta,time
from dateutil.relativedelta import relativedelta
from odoo.exceptions import UserError
from calendar import monthrange
from pytz import timezone, UTC
import io
import xlsxwriter
import base64
import logging
import calendar
import pytz
import logging

_logger = logging.getLogger(__name__)
class HrAttendance(models.Model):
    _inherit = 'hr.attendance'

    real_worked_hours = fields.Float(string='Horas reales', compute='_compute_real_hours')
    scheduled_check_in = fields.Datetime(string='Hora programada de entrada', compute='_compute_scheduled_attendance_times', store=True)

    @api.depends('check_in', 'check_out', 'scheduled_check_in')
    def _compute_real_hours(self):
        for record in self:
            if not record.check_in or not record.check_out:
                record.real_worked_hours = 0.0
                continue
    
            start = record.check_in
    
            if record.scheduled_check_in:
                start = max(record.check_in, record.scheduled_check_in)
    
            delta = record.check_out - start
            record.real_worked_hours = delta.total_seconds() / 3600.0


    @api.depends('employee_id', 'check_in')
    def _compute_scheduled_attendance_times(self):
        for attendance in self:
            if not attendance.employee_id or not attendance.check_in:
                attendance.scheduled_check_in = False
                #attendance.scheduled_check_out = False
                continue
    
            # Obtener contrato vigente en la fecha de check_in
            contract = self.env['hr.contract'].search([
                ('employee_id', '=', attendance.employee_id.id),
                ('state', '=', 'open'),
                ('date_start', '<=', attendance.check_in.date()),
                '|',
                ('date_end', '=', False),
                ('date_end', '>=', attendance.check_in.date())
            ], limit=1)
    
            if not contract or not contract.resource_calendar_id:
                attendance.scheduled_check_in = False
                #attendance.scheduled_check_out = False
                continue
    
            calendar = contract.resource_calendar_id
            employee = attendance.employee_id
            # ✅ Usar zona horaria de Paraguay por defecto si no está definida
            local_tz = timezone(employee.tz or 'America/Asuncion')
    
            # Convertir check_in a zona local para comparar con el calendario
            check_in_local = attendance.check_in.replace(tzinfo=UTC).astimezone(local_tz)
            day_start = local_tz.localize(datetime.combine(check_in_local.date(), time.min))
            day_end = local_tz.localize(datetime.combine(check_in_local.date(), time.max))

            candidates = []
            # Obtener intervalos laborales del día (excluye descansos automáticamente)
            intervals = employee._employee_attendance_intervals(
                day_start.astimezone(UTC),
                day_end.astimezone(UTC),
                lunch=False
            )
            normal_intervals = contract.resource_calendar_id._attendance_intervals_batch(
                day_start.astimezone(UTC),
                day_end.astimezone(UTC),
                employee.resource_id
            ).get(employee.resource_id.id, [])
            check_date = attendance.check_in.date()

            normal_intervals = sorted(normal_intervals, key=lambda x: x[0])

            for interval in normal_intervals:
                start_local = interval[0].astimezone(local_tz)
            
                if start_local.hour < 4:
                    continue
            
                candidates.append(interval)
                
            
            intervals_list = sorted(list(intervals), key=lambda x: x[0])
            
            # ✅ Obtener primer intervalo para entrada programada
            if intervals_list:
                # Buscar el intervalo más cercano al check_in
                check_in_local = attendance.check_in.replace(tzinfo=UTC).astimezone(local_tz)
                
                closest_interval = None
                smallest_diff = None
                for interval in candidates:
                    start = interval[0]
                    end = interval[1]
                
                    if start <= check_in_local <= end:
                        closest_interval = interval
                        break
                
                # fallback: usar el primer intervalo del día
                if not closest_interval and candidates:
                    closest_interval = candidates[0]
                
                if closest_interval:
                    scheduled_in_local = closest_interval[0]
                    scheduled_out_local = closest_interval[1]
                
                    attendance.scheduled_check_in = scheduled_in_local.astimezone(UTC).replace(tzinfo=None)
                    #attendance.scheduled_check_out = scheduled_out_local.astimezone(UTC).replace(tzinfo=None)
                else:
                    attendance.scheduled_check_in = False
                    #attendance.scheduled_check_out = False
            else:
                attendance.scheduled_check_in = False
    

                