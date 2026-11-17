"""Confirmed-input payroll arithmetic. No salary guessing, tax estimation or brokerage writes.
Amounts use Decimal, ROUND_HALF_UP at final item; policy ambiguities return pending.
"""
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from datetime import date
import calendar

def number(value):
    if value is None or str(value).strip()=='':raise ValueError('缺少已确认数字')
    try:value=Decimal(str(value))
    except InvalidOperation as exc:raise ValueError('已确认数字无法解析') from exc
    if not value.is_finite():raise ValueError('数字必须有限')
    return value
def money(value):return float(number(value).quantize(Decimal('.01'),rounding=ROUND_HALF_UP))
def personal_leave_deduction(salary,hours,intern,required_hours):
    denominator=number(required_hours) if intern else Decimal(174)
    if denominator<=0 or number(hours)<0:raise ValueError('工时分母须大于0；负结余另核')
    return money(-number(salary)/denominator*number(hours))
def maternity_difference(normal,maternity,workdays,maternity_days):
    if number(workdays)<=0 or not 0<=number(maternity_days)<=number(workdays):raise ValueError('产假工作日范围错误')
    return money((number(normal)-number(maternity))/number(workdays)*number(maternity_days))
def increment_amount(initial,regular,probation_months):
    if probation_months==3:return 0
    if probation_months!=6:raise ValueError('未定义的试用期递增口径')
    return money((number(regular)-number(initial))/5)
def cycle_boundary(hire,months):
    y,m=divmod(hire.year*12+hire.month-1+months,12);m+=1
    day=min(hire.day,calendar.monthrange(y,m)[1])
    if day!=hire.day:raise ValueError('周期对应日不存在，须确认月底周期边界')
    return date(y,m,day)
def prorate(amount,total_days,applicable_days):
    if number(total_days)<=0 or not 0<=number(applicable_days)<=number(total_days):raise ValueError('计薪天数范围错误')
    return money(number(amount)/number(total_days)*number(applicable_days))
def hire_exit_difference(salary,paid_days,unpaid_days):return -prorate(salary,paid_days,unpaid_days)
def sick_leave_deduction(salary,hours,required_hours,prior_hours,minimum_wage=None):
    try:
        h,prior,den=number(hours),number(prior_hours),number(required_hours)
        if den<=0 or min(h,prior)<0:raise ValueError('病假工时无效')
        if prior<72<prior+h:return dict(status='pending',amount=None,reason='本月跨年度9天阈值，分段制度尚未确认')
        if prior+h<=72:amount=-number(salary)/den*h/2
        else:amount=-(number(salary)-number(minimum_wage)*Decimal('.8'))/den*h
        return dict(status='calculated',amount=money(amount),reason='按已确认病假档位；最低工资须使用当月已核准来源')
    except ValueError as e:return dict(status='pending',amount=None,reason=str(e))
def payroll_totals(items):
    required=('fixed_salary','other_allowance','performance','reward','leave_deduction','attendance_deduction','hire_exit_difference','other','personal_deduction','personal_social','personal_medical','personal_fund','tax','combined_tax','third_party','cash_other')
    v={k:number(items.get(k)) for k in required}
    gross=sum(v[k] for k in required[:7]);finance=gross+v['other']
    net=finance-sum(v[k] for k in ('personal_deduction','personal_social','personal_medical','personal_fund','tax'))
    return dict(gross=money(gross),finance=money(finance),net=money(net),taxable=money(finance-v['personal_deduction']+v['combined_tax']),bank=money(net-v['third_party']-v['cash_other']))

def salary_progression(initial,regular,probation_months,hire,month,paid_dates,resignation=None):
    """Six-month confirmed offer, month-day calendar supplied by HR, no guessed holiday rules.
    Returns fixed salary at start-month plus only this month's partial increment allowance.
    Incomplete resignation cycle increments are excluded. Missing same-day boundaries stay pending.
    """
    try:
        hire=date.fromisoformat(str(hire));first=date.fromisoformat(month+'-01')
        last=date(first.year,first.month,calendar.monthrange(first.year,first.month)[1])
        resignation=date.fromisoformat(str(resignation)) if resignation else None
        days=[date.fromisoformat(str(d)) for d in paid_dates]
        if not days or len(days)!=len(set(days)) or any(not first<=d<=last for d in days):raise ValueError('计薪日历缺失、重复或跨月')
        step=number(increment_amount(initial,regular,probation_months));base=number(initial);allowance=Decimal(0);trace=[]
        if probation_months==3:return dict(status='calculated',salary=money(base),allowance=0,trace=['三个月试用期不递增；转正标准另依确认'])
        for k in range(1,6):
            effective=cycle_boundary(hire,k);end=cycle_boundary(hire,k+1)
            if effective>last:continue
            # Only an employee leaving before completion of this specific cycle loses its increment.
            if resignation and effective<=resignation<end:
                trace.append({'step':k,'effective':str(effective),'status':'excluded_incomplete_resignation_cycle'});continue
            if effective<=first:base+=step;trace.append({'step':k,'status':'fixed'})
            else:
                eligible=sum(1 for d in days if d>=effective and d>=hire and (not resignation or d<=resignation))
                allowance+=step/len(days)*eligible
                trace.append({'step':k,'status':'allowance','days':eligible,'denominator':len(days)})
        return dict(status='calculated',salary=money(base),allowance=money(allowance),trace=trace)
    except (ValueError,TypeError) as e:return dict(status='pending',salary=None,allowance=None,reason=str(e))
