from odoo import models, fields, api
from datetime import date,datetime,timedelta
from dateutil.relativedelta import relativedelta
from odoo.exceptions import UserError
from calendar import monthrange
import io
import xlsxwriter
import base64
import logging
import calendar
import pytz
import logging

_logger = logging.getLogger(__name__)
class InfoWizard(models.TransientModel):
    _name = 'hr.attendance.custom_report'
    _description = 'Reporte personalizado de asistencias'

    year = fields.Integer(string="Año", required=True, default=lambda self: fields.Date.today().year)
    company_id = fields.Many2one('res.company', string="Compañía", default=lambda self: self.env.company)
    month = fields.Selection([
        ('1', 'Enero'),
        ('2', 'Febrero'),
        ('3', 'Marzo'),
        ('4', 'Abril'),
        ('5', 'Mayo'),
        ('6', 'Junio'),
        ('7', 'Julio'),
        ('8', 'Agosto'),
        ('9', 'Septiembre'),
        ('10', 'Octubre'),
        ('11', 'Noviembre'),
        ('12', 'Diciembre'),
    ], string="Mes", required=True, default=lambda self: str(fields.Date.today().month))


    def generate_attendance_report_excel(self):
        self.ensure_one()
        company = self.company_id
        year = self.year
        month = int(self.month)
        num_days = monthrange(year, month)[1]
        start_date = date(year, month, 1)
        end_date = date(year, month, num_days)
        company_calendar = self.env.company.resource_calendar_id
        feriados = set()
        if company_calendar:
            leaves_feriados = self.env['resource.calendar.leaves'].search([
                ('calendar_id', '=', company_calendar.id),
                ('resource_id', '=', False),  # feriados globales
                ('date_to', '>=', fields.Datetime.to_datetime(start_date)),
                ('date_from', '<=', fields.Datetime.to_datetime(end_date + timedelta(days=1))),
            ])
            for leaf in leaves_feriados:
                from_date = max(leaf.date_from.date(), start_date)
                to_date = min(leaf.date_to.date(), end_date)
                current = from_date
                while current <= to_date:
                    feriados.add(current)
                    current += timedelta(days=1)

        # Buscar empleados activos de la compañía
        employees = self.env['hr.employee'].search([
            ('company_id', '=', company.id),
            ('active', '=', True)
        ])
    
        if not employees:
            raise UserError("No se encontraron empleados activos en la compañía %s." % company.name)
    
        # Crear archivo en memoria
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Asistencias')
    
        # Formatos
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#2C3E50',
            'font_color': 'white',
            'border': 1,
            'text_wrap': True
        })
        text_format = workbook.add_format({'border': 1, 'align': 'left'})
        number_format = workbook.add_format({'border': 1, 'align': 'right', 'num_format': '0.0'})
    
        # ✅ SOLO LAS 8 COLUMNAS REALES
        headers = [
            'SECTOR',
            'APELLIDO',
            'Jornada Recibo',
            'Lic x Enf',
            'Otras Lic',
            'Vacac',
            'SIN JUST.',
            'Present Quincena',
            'ILT',
            'Total',
        ]
    
        # Escribir cabeceras
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
    
        # Escribir datos por empleado
        row_index = 1
        for emp in employees:
            sector = emp.department_id.name or "SIN DEPARTAMENTO"
            last_name = emp.name or "Sin nombre"
            
            # Inicializar contadores
            sin_just_fechas = set() 
            horas_trabajadas = 0.0
            lic_enf = 0.0
            otras_lic = 0.0
            art = 0.0
            vacac = 0.0
            sin_just = 0.0
            
            # --- 1. Obtener todas las asistencias del mes (optimizado) ---
            attendances = self.env['hr.attendance'].search([
                ('employee_id', '=', emp.id),
                ('check_in', '>=', fields.Datetime.to_datetime(start_date)),
                ('check_in', '<=', fields.Datetime.to_datetime(end_date + timedelta(days=1))),
            ])
            
            # Agrupar asistencias por fecha (solo fecha, sin hora)
            attendance_by_date = {}
            for att in attendances:
                if not att.check_out:  # Saltear asistencias incompletas
                    continue
                att_date = att.check_in.date()
                if start_date <= att_date <= end_date:
                    if att_date not in attendance_by_date:
                        attendance_by_date[att_date] = []
                    attendance_by_date[att_date].append(att)
            
            horas_trabajadas = 0.0
            attendance_dates = set(attendance_by_date.keys())
            
            # Calendario del empleado
            calendar = emp.resource_calendar_id or self.env.company.resource_calendar_id
            
            for current_date, att_list in attendance_by_date.items():
                # Combinar todas las asistencias del día en un solo intervalo (o suma de intervalos)
                # Primero, ordenar por check_in
                att_list.sort(key=lambda x: x.check_in)
                
                # Fusionar solapamientos (opcional, pero seguro)
                merged = []
                for att in att_list:
                    if not merged:
                        merged.append([att.check_in, att.check_out])
                    else:
                        last = merged[-1]
                        if att.check_in <= last[1]:
                            last[1] = max(last[1], att.check_out)
                        else:
                            merged.append([att.check_in, att.check_out])
                
                total_worked_seconds = 0
                for interval in merged:
                    check_in_real = interval[0]
                    check_out_real = interval[1]
            
                    # --- Obtener horario contractual para este día ---
                    if calendar and emp.resource_id:
                        # Usar zona horaria
                        tz_name = self.env.user.tz or self.env.company.partner_id.tz or 'America/Asuncion'
                        local_tz = pytz.timezone(tz_name)
                        
                        # Convertir fecha a datetime en zona local
                        start_dt = local_tz.localize(datetime.combine(current_date, datetime.min.time()))
                        end_dt = start_dt + timedelta(days=1)
            
                        # Obtener intervalos laborales del calendario
                        work_intervals = calendar._work_intervals_batch(start_dt, end_dt, resources=emp.resource_id)
                        intervals_for_day = list(work_intervals.get(emp.resource_id.id, []))
                        
                        if intervals_for_day:
                            # Tomar el primer intervalo como horario principal (puedes ajustar si hay varios turnos)
                            first_interval = intervals_for_day[0]
                            scheduled_start = first_interval[0]  # datetime con tz
                            scheduled_end = first_interval[1]
                            
                            # Convertir asistencias a la misma zona horaria
                            check_in_tz = check_in_real.astimezone(local_tz)
                            check_out_tz = check_out_real.astimezone(local_tz)
                            _logger.info(f"Empleado:{emp.name}")
                            _logger.info(check_in_tz)
                            _logger.info(check_out_tz)
                            # Ajustar entrada: no antes del horario
                            # Ajustar entrada: no antes del horario contractual
                            actual_start = max(check_in_tz, scheduled_start)
                            # ✅ NO limitar la salida: permitir que se quede más si entró tarde
                            actual_end = check_out_tz
                            
                            if actual_end > actual_start:
                                worked = (actual_end - actual_start).total_seconds()
                                _logger.info(worked)
                                total_worked_seconds += worked
                        else:
                            # Día no laborable: no cuenta (o podrías ignorar)
                            continue
                    else:
                        # Sin calendario: usar las horas reales, pero máximo 9h
                        worked = (check_out_real - check_in_real).total_seconds()
                        total_worked_seconds += worked
            
                # Limitar a 9 horas (32400 segundos)
                total_worked_seconds = min(total_worked_seconds, 9 * 3600)
                horas_del_dia = total_worked_seconds / 3600
                horas_trabajadas += round(horas_del_dia * 2) / 2  # redondear a 0.5
            # --- 2. Obtener todos los permisos validados del mes ---
            leaves = self.env['hr.leave'].search([
                ('employee_id', '=', emp.id),
                ('state', '=', 'validate'),
                ('date_to', '>=', fields.Datetime.to_datetime(start_date)),  # termina después del inicio del mes
                ('date_from', '<=', fields.Datetime.to_datetime(end_date + timedelta(days=1))),  # empieza antes del día siguiente al fin del mes
            ])
                        
            covered_dates = set()
            
            for leave in leaves:
                # Determinar rango real del permiso dentro del mes
                from_date = max(leave.date_from.date(), start_date)
                to_date = min(leave.date_to.date(), end_date)
                if from_date > to_date:
                    continue
            
                # Iterar día por día
                current = from_date
                while current <= to_date:
                    # Saltar si ya fue cubierto (evitar superposiciones)
                    if current in covered_dates or current in attendance_dates:
                        current += timedelta(days=1)
                        continue
            
                    # Verificar si es día laborable (igual que number_of_days)
                    calendar = emp.resource_calendar_id or self.env.company.resource_calendar_id
                    is_work_day = True
                    if calendar and emp.resource_id:
                        # Convertir la fecha a datetime con zona horaria
                        tz = self.env.user.tz or self.env.company.partner_id.tz or 'America/Asuncion'
                        from pytz import timezone
                        local_tz = timezone(tz)
                    
                        start_naive = datetime.combine(current, datetime.min.time())
                        start_local = local_tz.localize(start_naive, is_dst=None)
                        end_local = start_local + timedelta(days=1)
                    
                        intervals = calendar._work_intervals_batch(start_local, end_local, resources=emp.resource_id)
                        is_work_day = bool(intervals.get(emp.resource_id.id))
                    if is_work_day:
                        covered_dates.add(current)
                        leave_name = (leave.holiday_status_id.name or '').lower()
                        if emp.name == 'Brusa María Mabel':
                            _logger.info(leave_name)
                        if 'enfermedad' in leave_name:
                            lic_enf += 9.0
                        elif 'vacaci' in leave_name:
                            vacac += 9.0
                        elif 'sin just' in leave_name or 'no just' in leave_name or 'ausencia por' in leave_name or 'ausencia por' in leave_name:
                            sin_just += 9.0
                            sin_just_fechas.add(current)
                        elif 'art' in leave_name:
                            art += 9.0
                        else:
                            otras_lic += 9.0
            
                    current += timedelta(days=1)
            # --- 3. Ahora, "Present Quincena" solo debe considerar días con asistencia ---#
            _logger.info(sin_just_fechas)
            q1_sin_just = any(
                date(year, month, day) in sin_just_fechas
                for day in range(1, min(16, num_days + 1))
            )
            
            q2_sin_just = any(
                date(year, month, day) in sin_just_fechas
                for day in range(16, num_days + 1)
            ) if num_days >= 16 else False
            
            # Aplicar regla: ausencia injustificada = falta en esa quincena
            if not q1_sin_just and not q2_sin_just:
                present_quincena = 1.0
            elif not q1_sin_just or not q2_sin_just:
                present_quincena = 0.5
            else:
                present_quincena = 0.0
    
            # ✅ Solo los 8 valores reales
            total_horas = horas_trabajadas + lic_enf + otras_lic + vacac + sin_just + art
            row_data = [
                sector,             # SECTOR
                last_name,          # APELLIDO
                horas_trabajadas,   # Jornada Recibo
                lic_enf,            # Lic x Enf
                otras_lic,          # Otras Lic
                vacac,              # Vacac
                sin_just,           # SIN JUST.
                present_quincena,   # Present Quincena
                art,
                total_horas
            ]
    
            # Escribir fila
            for col, value in enumerate(row_data):
                if isinstance(value, (int, float)):
                    worksheet.write(row_index, col, value, number_format)
                else:
                    worksheet.write(row_index, col, value or '', text_format)
            row_index += 1
    
        # Ajustar ancho de columnas
        worksheet.set_column('A:H', 20)
    
        workbook.close()
        output.seek(0)
        file_data = base64.b64encode(output.read())
        output.close()
    
        filename = f"Asistencia_{company.name}_{year}_{month:02d}.xlsx"
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'type': 'binary',
            'datas': file_data,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
    
        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'self',
        }


