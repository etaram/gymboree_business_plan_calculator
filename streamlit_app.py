import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf
import plotly.express as px
import plotly.graph_objects as go
import locale
import io

# הגדרת לוקאל לאלפי מפרידים
try:
    locale.setlocale(locale.LC_ALL, 'he_IL.UTF-8')
except locale.Error:
    locale.setlocale(locale.LC_ALL, '')

# כותרת האפליקציה
st.title("מחשבון תוכנית עסקית לג'ימבורי יובל המבולבל")

# פרמטרים ראשוניים
default_params = {
    # עלויות הקמה
    'עלות בנייה': 900_000,
    'עלות מערכות חשמל ותאורה': 300_000,
    'עלות מערכות מיזוג ואוורור': 250_000,
    'עלות מתקני משחק': 900_000,
    'עלות ציוד VR/AR': 300_000,
    'עלות ריהוט ואביזרים': 150_000,
    'עלות מערכות ניהול ובקרה': 150_000,
    'עלות מערכות קופות ותשלומים': 100_000,
    'עלות אתר אינטרנט': 50_000,
    'עלות אישורי בטיחות וכיבוי אש': 100_000,
    'עלות רישיונות עסק': 50_000,
    'עלות ייעוץ עסקי ופיננסי': 50_000,
    'הוצאות משפטיות': 50_000,
    'הון חוזר לתפעול ראשוני': 500_000,

    # פרמטרים תפעוליים
    'מחיר כניסה ליום רגיל': 50,
    'מספר מבקרים ביום רגיל': 300,
    'מחיר כניסה ליום חופשה/חג': 50,
    'מספר מבקרים ביום חופשה/חג': 1_200,
    'רכישה ממוצעת במזון ביום רגיל': 20,
    'רכישה ממוצעת במזון ביום חופשה/חג': 25,
    'רכישה ממוצעת במרצ\'נדייז ביום רגיל': 15,
    'רכישה ממוצעת במרצ\'נדייז ביום חופשה/חג': 20,
    'מספר אירועים פרטיים בחודש': 20,
    'מחיר לאירוע פרטי': 2_000,
    'מספר סדנאות בחודש': 5,
    'מספר משתתפים בסדנה': 15,
    'מחיר לסדנה': 100,

    # הוצאות קבועות
    'שכר דירה חודשי': 150_000,
    'משכורת מנכ"ל': 20_000,
    'משכורת מנהלים (סה"כ)': 60_000,
    'משכורת צוות (סה"כ)': 140_000,
    'ארנונה שנתית': 360_000,
    'הוצאות חשמל חודשיות': 25_000,
    'הוצאות מים חודשיות': 5_000,
    'הוצאות נוספות שנתיות': 236_000,
    'תשלומי הלוואה שנתיים': 975_000,

    # פרמטרים נוספים
    'שיעור פחת שנתי (%)': 2,
    'אורך מימון (שנים)': 5
}

# פונקציה לחישוב עלויות הקמה
def calculate_setup_costs(params):
    return sum([
        params['עלות בנייה'], params['עלות מערכות חשמל ותאורה'], params['עלות מערכות מיזוג ואוורור'],
        params['עלות מתקני משחק'], params['עלות ציוד VR/AR'], params['עלות ריהוט ואביזרים'],
        params['עלות מערכות ניהול ובקרה'], params['עלות מערכות קופות ותשלומים'], params['עלות אתר אינטרנט'],
        params['עלות אישורי בטיחות וכיבוי אש'], params['עלות רישיונות עסק'], params['עלות ייעוץ עסקי ופיננסי'],
        params['הוצאות משפטיות']
    ])

# פונקציה לחישוב הכנסות
def calculate_income(params):
    income_regular_tickets = params['מספר מבקרים ביום רגיל'] * params['מחיר כניסה ליום רגיל'] * 191
    income_regular_food = params['מספר מבקרים ביום רגיל'] * params['רכישה ממוצעת במזון ביום רגיל'] * 191
    income_regular_merch = params['מספר מבקרים ביום רגיל'] * params['רכישה ממוצעת במרצ\'נדייז ביום רגיל'] * 191

    income_holiday_tickets = params['מספר מבקרים ביום חופשה/חג'] * params['מחיר כניסה ליום חופשה/חג'] * 100
    income_holiday_food = params['מספר מבקרים ביום חופשה/חג'] * params['רכישה ממוצעת במזון ביום חופשה/חג'] * 100
    income_holiday_merch = params['מספר מבקרים ביום חופשה/חג'] * params['רכישה ממוצעת במרצ\'נדייז ביום חופשה/חג'] * 100

    income_events = params['מספר אירועים פרטיים בחודש'] * params['מחיר לאירוע פרטי'] * 12
    income_workshops = params['מספר סדנאות בחודש'] * params['מספר משתתפים בסדנה'] * params['מחיר לסדנה'] * 12

    return sum([
        income_regular_tickets, income_regular_food, income_regular_merch,
        income_holiday_tickets, income_holiday_food, income_holiday_merch,
        income_events, income_workshops
    ])

# פונקציה לחישוב הוצאות משתנות
def calculate_variable_expenses(params, income_regular_merch, income_holiday_merch, income_regular_food, income_holiday_food, income_events, income_workshops):
    cost_of_merch = (income_regular_merch + income_holiday_merch) * 0.5
    cost_of_food = (income_regular_food + income_holiday_food) * 0.4
    cost_of_workshops = income_workshops * 0.5
    cost_of_events = income_events * 0.4

    return cost_of_merch + cost_of_food + cost_of_workshops + cost_of_events

# פונקציה לחישוב הוצאות קבועות
def calculate_fixed_expenses(params, annual_depreciation):
    return (
        (params['שכר דירה חודשי'] + params['משכורת מנכ"ל'] + params['משכורת מנהלים (סה"כ)'] +
         params['משכורת צוות (סה"כ)'] + params['הוצאות חשמל חודשיות'] + params['הוצאות מים חודשיות']) * 12 +
        params['ארנונה שנתית'] + params['הוצאות נוספות שנתיות'] + annual_depreciation + params['הון חוזר לתפעול ראשוני']
    )

# פונקציה לחישוב תשלומי הלוואה
def calculate_loan_payments(params):
    loan_duration_years = int(params['אורך מימון (שנים)'])
    loan_duration_months = loan_duration_years * 12
    loan_amount = params['תשלומי הלוואה שנתיים']
    annual_interest_rate = 0.05 
    monthly_interest_rate = annual_interest_rate / 12
    if monthly_interest_rate == 0:
        monthly_payment = loan_amount / loan_duration_months
    else:
        monthly_payment = loan_amount * monthly_interest_rate / (1 - (1 + monthly_interest_rate) ** (-loan_duration_months))
    return monthly_payment * loan_duration_months

# פונקציה לחישוב רווח לפני מס
def calculate_profit_before_tax(gross_profit, total_fixed_expenses, total_loan_payments):
    return gross_profit - total_fixed_expenses - total_loan_payments

# פונקציה לחישוב מחיר כניסה ממוצע
def calculate_average_ticket_price(total_income, params):
    total_visitors_regular = params['מספר מבקרים ביום רגיל'] * 191
    total_visitors_holiday = params['מספר מבקרים ביום חופשה/חג'] * 100
    total_visitors = total_visitors_regular + total_visitors_holiday

    return total_income / total_visitors if total_visitors != 0 else 0

# פונקציה לחישוב נקודת איזון כוללת (מספר כרטיסים לשנה)
def calculate_breakeven_total_tickets(total_variable_expenses, total_fixed_expenses, total_loan_payments, average_ticket_price):
    return (total_variable_expenses + total_fixed_expenses + total_loan_payments) / average_ticket_price if average_ticket_price != 0 else 0

# פונקציה לחישוב נקודת איזון יומית
def calculate_breakeven_daily_tickets(breakeven_total_tickets):
    operating_days_per_year = 300 
    return breakeven_total_tickets / operating_days_per_year if operating_days_per_year != 0 else 0

# פונקציה לחישוב החזר על ההשקעה (ROI)
def calculate_roi(profit_before_tax, setup_costs):
    return (profit_before_tax / setup_costs) * 100 if setup_costs != 0 else 0

# פונקציה לחישוב תקופת החזר ההשקעה
def calculate_payback_period(setup_costs, profit_before_tax):
    return setup_costs / profit_before_tax if profit_before_tax != 0 else 0

# פונקציה לחישוב החזר פנימי (IRR)
def calculate_irr(setup_costs, profit_before_tax, loan_duration_years):
    cash_flows = [-setup_costs] + [profit_before_tax] * loan_duration_years
    try:
        return npf.irr(cash_flows) * 100 
    except:
        return None

# פונקציה לחישוב הכנסות והוצאות לפי קטגוריות
def calculate_category_data(params, income_regular_tickets, income_holiday_tickets, income_regular_food, income_holiday_food,
                             income_regular_merch, income_holiday_merch, income_events, income_workshops,
                             cost_of_merch, cost_of_food, cost_of_events, cost_of_workshops, total_fixed_expenses, total_loan_payments):
    income_data = {
        'כרטיסים': income_regular_tickets + income_holiday_tickets,
        'מזון ומשקאות': income_regular_food + income_holiday_food,
        'מרצ\'נדייז': income_regular_merch + income_holiday_merch,
        'אירועים': income_events,
        'סדנאות': income_workshops
    }

    expense_data = {
        'הוצאות קבועות': total_fixed_expenses,
        'עלות מרצ\'נדייז': cost_of_merch,
        'עלות מזון ומשקאות': cost_of_food,
        'עלות אירועים': cost_of_events,
        'עלות סדנאות': cost_of_workshops,
        'תשלומי הלוואה': total_loan_payments
    }

    return income_data, expense_data


# פונקציה לחישוב התוצאות
def calculate_results(params):
    """
    מחשב את התוצאות הפיננסיות של התוכנית העסקית בהתבסס על פרמטרי הקלט.

    Args:
        params (dict): מילון של פרמטרים.

    Returns:
        dict: מילון של תוצאות.
    """

    setup_costs = calculate_setup_costs(params)

    # פחת שנתי כאחוז
    annual_depreciation_percentage = params['שיעור פחת שנתי (%)']
    annual_depreciation = setup_costs * (annual_depreciation_percentage / 100)

    # הכנסות
    income_regular_tickets = params['מספר מבקרים ביום רגיל'] * params['מחיר כניסה ליום רגיל'] * 191
    income_regular_food = params['מספר מבקרים ביום רגיל'] * params['רכישה ממוצעת במזון ביום רגיל'] * 191
    income_regular_merch = params['מספר מבקרים ביום רגיל'] * params['רכישה ממוצעת במרצ\'נדייז ביום רגיל'] * 191

    income_holiday_tickets = params['מספר מבקרים ביום חופשה/חג'] * params['מחיר כניסה ליום חופשה/חג'] * 100
    income_holiday_food = params['מספר מבקרים ביום חופשה/חג'] * params['רכישה ממוצעת במזון ביום חופשה/חג'] * 100
    income_holiday_merch = params['מספר מבקרים ביום חופשה/חג'] * params['רכישה ממוצעת במרצ\'נדייז ביום חופשה/חג'] * 100

    income_events = params['מספר אירועים פרטיים בחודש'] * params['מחיר לאירוע פרטי'] * 12
    income_workshops = params['מספר סדנאות בחודש'] * params['מספר משתתפים בסדנה'] * params['מחיר לסדנה'] * 12

    total_income = calculate_income(params)

    # הוצאות משתנות
    total_variable_expenses = calculate_variable_expenses(
        params, income_regular_merch, income_holiday_merch, income_regular_food, income_holiday_food, income_events, income_workshops
    )

    # רווח גולמי
    gross_profit = total_income - total_variable_expenses

    # הוצאות קבועות (כוללות פחת והון חוזר)
    total_fixed_expenses = calculate_fixed_expenses(params, annual_depreciation)

    # תשלומי הלוואה חודשיים
    total_loan_payments = calculate_loan_payments(params)

    # רווח לפני מס
    profit_before_tax = calculate_profit_before_tax(gross_profit, total_fixed_expenses, total_loan_payments)

    # מחיר כניסה ממוצע
    average_ticket_price = calculate_average_ticket_price(total_income, params)

    # נקודת איזון כוללת (מספר כרטיסים לשנה)
    breakeven_total_tickets = calculate_breakeven_total_tickets(total_variable_expenses, total_fixed_expenses, total_loan_payments, average_ticket_price)

    # נקודת איזון יומית
    breakeven_daily_tickets = calculate_breakeven_daily_tickets(breakeven_total_tickets)

    # החזר על ההשקעה (ROI)
    roi = calculate_roi(profit_before_tax, setup_costs)

    # תקופת החזר ההשקעה
    payback_period = calculate_payback_period(setup_costs, profit_before_tax)

    # החזר פנימי (IRR)
    irr = calculate_irr(setup_costs, profit_before_tax, int(params['אורך מימון (שנים)']))

    # הכנסות והוצאות לפי קטגוריות
    income_data, expense_data = calculate_category_data(
        params, income_regular_tickets, income_holiday_tickets, income_regular_food, income_holiday_food,
        income_regular_merch, income_holiday_merch, income_events, income_workshops,
        total_variable_expenses * (
            (income_regular_merch + income_holiday_merch) * 0.5 / total_variable_expenses if total_variable_expenses != 0 else 0),
        total_variable_expenses * (
            (income_regular_food + income_holiday_food) * 0.4 / total_variable_expenses if total_variable_expenses != 0 else 0),
        total_variable_expenses * (income_events * 0.4 / total_variable_expenses if total_variable_expenses != 0 else 0),
        total_variable_expenses * (
            income_workshops * 0.5 / total_variable_expenses if total_variable_expenses != 0 else 0),
        total_fixed_expenses, total_loan_payments
    )

    results = {
        'עלויות הקמה': setup_costs,
        'פחת שנתי (%)': annual_depreciation_percentage, 
        'הכנסות שנתיות': total_income,
        'הוצאות משתנות': total_variable_expenses,
        'רווח גולמי': gross_profit,
        'הוצאות קבועות': total_fixed_expenses,
        'תשלומי הלוואה': total_loan_payments,
        'רווח לפני מס': profit_before_tax,
        'נקודת איזון (מספר כרטיסים לשנה)': breakeven_total_tickets,
        'נקודת איזון (מספר כרטיסים ליום)': breakeven_daily_tickets,
        'החזר על ההשקעה (ROI)': roi,
        'החזר פנימי (IRR)': irr,
        'תקופת החזר השקעה (שנים)': payback_period,
        'הכנסות לפי קטגוריות': income_data,
        'הוצאות לפי קטגוריות': expense_data
    }

    return results

# יצירת שדה קלט (מותאם ל-Streamlit)
def create_input_control(key, value):
    if isinstance(value, int):
        formatted_value = f"{locale.format_string('%d', value, grouping=True)}"
    elif isinstance(value, float):
        formatted_value = f"{locale.format_string('%.2f', value, grouping=True)}"
    else:
        formatted_value = value

    if 'שנים' in key or 'מספר' in key or 'שיעור' in key:
        input_type = 'number'
    else:
        input_type = 'text'

    if 'שיעור פחת שנתי (%)' in key:
        suffix = '%'
    elif 'שיעור' in key:
        suffix = '%'
    elif 'עלות' in key or 'הוצאות' in key or 'משכורת' in key or 'תשלומים' in key or 'רכישה' in key or 'מחיר' in key:
        suffix = ' ש״ח'
    else:
        suffix = ''

    if input_type == 'number':
        return st.number_input(f"{key}{suffix}", value=value)
    else:
        return st.text_input(f"{key}{suffix}", value=formatted_value)


# הצגת טאב הפרמטרים (מותאם ל-Streamlit)
def render_parameters_tab():
    st.header("פרמטרים")
    
    # קבוצות פרמטרים
    setup_params = [
        'עלות בנייה', 'עלות מערכות חשמל ותאורה', 'עלות מערכות מיזוג ואוורור', 'עלות מתקני משחק', 'עלות ציוד VR/AR',
        'עלות ריהוט ואביזרים', 'עלות מערכות ניהול ובקרה', 'עלות מערכות קופות ותשלומים', 'עלות אתר אינטרנט',
        'עלות אישורי בטיחות וכיבוי אש', 'עלות רישיונות עסק', 'עלות ייעוץ עסקי ופיננסי', 'הוצאות משפטיות', 'הון חוזר לתפעול ראשוני',
        'שיעור פחת שנתי (%)'
    ]

    operational_params = [
        'מחיר כניסה ליום רגיל', 'מספר מבקרים ביום רגיל', 'מחיר כניסה ליום חופשה/חג', 'מספר מבקרים ביום חופשה/חג',
        'רכישה ממוצעת במזון ביום רגיל', 'רכישה ממוצעת במזון ביום חופשה/חג', 'רכישה ממוצעת במרצ\'נדייז ביום רגיל',
        'רכישה ממוצעת במרצ\'נדייז ביום חופשה/חג', 'מספר אירועים פרטיים בחודש', 'מחיר לאירוע פרטי',
        'מספר סדנאות בחודש', 'מספר משתתפים בסדנה', 'מחיר לסדנה'
    ]

    fixed_expenses_params = [
        'שכר דירה חודשי', 'משכורת מנכ"ל', 'משכורת מנהלים (סה"כ)', 'משכורת צוות (סה"כ)',
        'ארנונה שנתית', 'הוצאות חשמל חודשיות', 'הוצאות מים חודשיות', 'הוצאות נוספות שנתיות', 'תשלומי הלוואה שנתיים',
        'אורך מימון (שנים)'
    ]

    with st.expander("עלויות הקמה"):
        for key in setup_params:
            default_params[key] = create_input_control(key, default_params[key])

    with st.expander("פרמטרים תפעוליים"):
        for key in operational_params:
            default_params[key] = create_input_control(key, default_params[key])

    with st.expander("הוצאות קבועות"):
        for key in fixed_expenses_params:
            default_params[key] = create_input_control(key, default_params[key])

    if st.button("חשב"):
        results = calculate_results(default_params)
        st.session_state['results'] = results

# פונקציה ליצירת טבלת מדדים
def generate_metrics_table(metrics, explanations):
    """
    יוצר טבלה של מדדים עם הסברים.

    Args:
        metrics (dict): מילון של מדדים.
        explanations (dict): מילון של הסברים למדדים.

    Returns:
        pd.DataFrame: טבלת המדדים.
    """

    metrics_df = pd.DataFrame([
        {'מדד': key, 'ערך': value, 'הסבר': explanations[key]}
        for key, value in metrics.items()
    ], columns=['מדד', 'ערך', 'הסבר'])

    return metrics_df

# פונקציה ליצירת טבלת קטגוריות
def generate_category_table(data):
    """
    יוצר טבלה של קטגוריות עם עיצוב אחיד.

    Args:
        data (dict): מילון של קטגוריות וסכומים.

    Returns:
        pd.DataFrame: טבלת הקטגוריות.
    """

    df = pd.DataFrame(list(data.items()), columns=['קטגוריה', 'סכום'])
    df['סכום'] = df['סכום'].apply(lambda x: f"{locale.format_string('%.2f', x, grouping=True)} ש״ח")  # פורמט סכום

    return df

# הצגת טאב התוצאות (מותאם ל-Streamlit)
def render_results_tab(results):
    if not results:
        st.write("אנא הזן את הפרמטרים ולחץ על 'חשב' כדי לראות את התוצאות.")
        return

    # הגדרת סף לממוצע כדי לקבוע צבעים
    roi_thresholds = {'high': 20, 'medium': 10}  # מעל 20% ירוק, 10-20% צהוב, מתחת 10% אדום
    irr_thresholds = {'high': 15, 'medium': 8}  # מעל 15% ירוק, 8-15% צהוב, מתחת 8% אדום

    def get_color(metric, value):
        if metric == 'ROI':
            if value >= roi_thresholds['high']:
                return 'success'  # ירוק
            elif value >= roi_thresholds['medium']:
                return 'warning'  # צהוב
            else:
                return 'danger'  # אדום
        elif metric == 'IRR':
            if value is None:
                return 'secondary'
            if value >= irr_thresholds['high']:
                return 'success'
            elif value >= irr_thresholds['medium']:
                return 'warning'
            else:
                return 'danger'
        else:
            return 'light'

    # יצירת טבלה של המדדים המרכזיים עם הסברים
    main_metrics = {
        'רווח לפני מס': f"{locale.format_string('%.2f', results['רווח לפני מס'], grouping=True)} ש״ח",
        'נקודת איזון (מספר כרטיסים לשנה)': f"{locale.format_string('%.2f', results['נקודת איזון (מספר כרטיסים לשנה)'], grouping=True)} כרטיסים לשנה",
        'נקודת איזון (מספר כרטיסים ליום)': f"{locale.format_string('%.2f', results['נקודת איזון (מספר כרטיסים ליום)'], grouping=True)} כרטיסים ליום",
        'החזר על ההשקעה (ROI)': f"{locale.format_string('%.2f', results['החזר על ההשקעה (ROI)'], grouping=True)}%",
        'החזר פנימי (IRR)': f"{locale.format_string('%.2f', results['החזר פנימי (IRR)'], grouping=True)}%" if results[
                                                                                                                 'החזר פנימי (IRR)'] is not None else "N/A",
        'תקופת החזר השקעה (שנים)': f"{locale.format_string('%.2f', results['תקופת החזר השקעה (שנים)'], grouping=True)} שנים"
    }

    # הסברים לכל מדד
    metrics_explanations = {
        'רווח לפני מס': "הרווח שנשאר לאחר ניכוי כל ההוצאות, כולל הוצאות קבועות ומשכורות, אך לפני ניכוי מס.",
        'נקודת איזון (מספר כרטיסים לשנה)': "מספר הכרטיסים שנדרש למכור בשנה כדי לכסות את כל ההכנסות וההוצאות ללא רווח או הפסד.",
        'נקודת איזון (מספר כרטיסים ליום)': "מספר הכרטיסים שנדרש למכור ביום פתוח כדי לכסות את ההוצאות.",
        'החזר על ההשקעה (ROI)': "החזר על ההשקעה מחושב כאחוז מהרווח לפני מס חלקי עלויות ההקמה.",
        'החזר פנימי (IRR)': "שיעור התשואה הפנימי שמאפשר לנקות את הערך הנוכחי של זרמי המזומנים להשקעה.",
        'תקופת החזר השקעה (שנים)': "הזמן שנדרש כדי להחזיר את ההשקעה הראשונית מהרווח לפני מס."
    }

    # יצירת טבלת המדדים המרכזיים עם הסברים
    metrics_table = generate_metrics_table(main_metrics, metrics_explanations)

    # יצירת טבלה של כל ההכנסות וההוצאות עם הסברים
    financials = {
        'הכנסות שנתיות': f"{locale.format_string('%.2f', results['הכנסות שנתיות'], grouping=True)} ש״ח",
        'הוצאות משתנות': f"{locale.format_string('%.2f', results['הוצאות משתנות'], grouping=True)} ש״ח",
        'רווח גולמי': f"{locale.format_string('%.2f', results['רווח גולמי'], grouping=True)} ש״ח",
        'הוצאות קבועות': f"{locale.format_string('%.2f', results['הוצאות קבועות'], grouping=True)} ש״ח",
        'תשלומי הלוואה': f"{locale.format_string('%.2f', results['תשלומי הלוואה'], grouping=True)} ש״ח",
        'רווח לפני מס': f"{locale.format_string('%.2f', results['רווח לפני מס'], grouping=True)} ש״ח"
    }

    # הסברים לכל מדד כלכלי
    financials_explanations = {
        'הכנסות שנתיות': "סה\"כ ההכנסות שהעסק צפוי לייצר בשנה.",
        'הוצאות משתנות': "ההוצאות התלויות בהכנסות, כמו עלויות מזון ומרצ\'נדייז.",
        'רווח גולמי': "הרווח לאחר ניכוי ההוצאות המשתנות מההכנסות.",
        'הוצאות קבועות': "הוצאות שאינן תלויות ישירות בהכנסות, כמו שכר דירה ומשכורות.",
        'תשלומי הלוואה': "סה\"כ התשלומים החודשיים לשירות ההלוואה.",
        'רווח לפני מס': "הרווח שנשאר לאחר ניכוי כל ההוצאות, כולל הוצאות קבועות ותשלומי הלוואה, אך לפני ניכוי מס."
    }

    financials_table = generate_metrics_table(financials, financials_explanations)
       # הכנסות והוצאות לפי קטגוריות עם הסברים
    income_table = generate_category_table(results['הכנסות לפי קטגוריות'])
    expense_table = generate_category_table(results['הוצאות לפי קטגוריות'])

    # יצירת מפתח צבעים למדדים העסקיים המרכזיים
    st.subheader("מפתח צבעים")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.write(
            "<span style='color:green;'>ROI גבוה (>=20%)</span>",
            unsafe_allow_html=True
        )
    with col2:
        st.write(
            "<span style='color:orange;'>ROI בינוני (10%-20%)</span>",
            unsafe_allow_html=True
        )
    with col3:
        st.write(
            "<span style='color:red;'>ROI נמוך (<10%)</span>",
            unsafe_allow_html=True
        )

    col1, col2, col3 = st.columns(3)
    with col1:
        st.write(
            "<span style='color:green;'>IRR גבוה (>=15%)</span>",
            unsafe_allow_html=True
        )
    with col2:
        st.write(
            "<span style='color:orange;'>IRR בינוני (8%-15%)</span>",
            unsafe_allow_html=True
        )
    with col3:
        st.write(
            "<span style='color:red;'>IRR נמוך (<8%)</span>",
            unsafe_allow_html=True
        )
    st.markdown("---")

    # יצירת טבלת המדדים המרכזיים עם הסברים
    st.subheader("מדדים עסקיים מרכזיים")
    st.write(metrics_table)
    st.markdown("---")

    # יצירת טבלת כל ההכנסות וההוצאות עם הסברים
    st.subheader("תוצאות כלכליות")
    st.write(financials_table)
    st.markdown("---")

    # הכנסות והוצאות לפי קטגוריות עם הסברים
    st.subheader("הכנסות לפי קטגוריות")
    st.write(income_table)
    st.markdown("---")

    st.subheader("הוצאות לפי קטגוריות")
    st.write(expense_table)
    st.markdown("---")


# הצגת טאב הגרפים (מותאם ל-Streamlit)
def render_charts_tab(results):
    """
    מציג את טאב הגרפים.

    Args:
        results (dict): תוצאות החישוב.

    Returns:
        html.Div: טאב הגרפים.
    """

    if not results:
        st.write("אנא הזן את הפרמטרים ולחץ על 'חשב' כדי לראות את הגרפים.")
        return

    # גרף מפל
    waterfall_fig = go.Figure(go.Waterfall(
        name="",
        orientation="v",
        measure=["relative", "relative", "relative", "relative", "total"],
        x=["הכנסות", "הוצאות משתנות", "הוצאות קבועות", "תשלומי הלוואה", "רווח לפני מס"],
        textposition="outside",
        text=[
            f"{locale.format_string('%.0f', results['הכנסות שנתיות'], grouping=True)} ש״ח",
            f"-{locale.format_string('%.0f', results['הוצאות משתנות'], grouping=True)} ש״ח",
            f"-{locale.format_string('%.0f', results['הוצאות קבועות'], grouping=True)} ש״ח",
            f"-{locale.format_string('%.0f', results['תשלומי הלוואה'], grouping=True)} ש״ח",
            f"{locale.format_string('%.0f', results['רווח לפני מס'], grouping=True)} ש״ח"
        ],
        y=[
            results['הכנסות שנתיות'],
            -results['הוצאות משתנות'],
            -results['הוצאות קבועות'],
            -results['תשלומי הלוואה'],
            results['רווח לפני מס']
        ],
        connector={"line": {"color": "rgb(63, 63, 63)"}},
        increasing={"marker": {"color": "#007bff"}},
        decreasing={"marker": {"color": "#dc3545"}},
        totals={"marker": {"color": "#28a745"}},
    ))
    waterfall_fig.update_layout(
        title="מפל הכנסות והוצאות",
        xaxis_title="",
        yaxis_title="ש״ח",
        font=dict(size=14),
        hovermode="x unified"
    )

    # גרף הכנסות לפי קטגוריות (גרף עוגה)
    income_df = pd.DataFrame(list(results['הכנסות לפי קטגוריות'].items()), columns=['קטגוריה', 'סכום'])
    income_pie_fig = px.pie(income_df, values='סכום', names='קטגוריה', title='התפלגות ההכנסות')
    income_pie_fig.update_traces(marker=dict(colors=['#007bff', '#28a745', '#dc3545', '#28a745', '#dc3545']))
    income_pie_fig.update_layout(font=dict(size=14))

    # גרף הוצאות לפי קטגוריות (גרף עוגה)
    expense_df = pd.DataFrame(list(results['הוצאות לפי קטגוריות'].items()), columns=['קטגוריה', 'סכום'])
    expense_pie_fig = px.pie(expense_df, values='סכום', names='קטגוריה', title='התפלגות ההוצאות')
    expense_pie_fig.update_layout(font=dict(size=14))

    # גרף נקודת איזון
    breakeven_total_tickets = results['נקודת איזון (מספר כרטיסים לשנה)']
    average_ticket_price = results['הכנסות שנתיות'] / (
            default_params['מספר מבקרים ביום רגיל'] * 191 + default_params['מספר מבקרים ביום חופשה/חג'] * 100) if (
            default_params['מספר מבקרים ביום רגיל'] * 191 + default_params['מספר מבקרים ביום חופשה/חג'] * 100) != 0 else 0

    quantity = np.linspace(0, breakeven_total_tickets * 1.5, 100)
    total_revenue = average_ticket_price * quantity
    if breakeven_total_tickets != 0:
        total_costs = (results['הוצאות משתנות'] + results['הוצאות קבועות'] + results['תשלומי הלוואה']) * quantity / breakeven_total_tickets
    else:
        total_costs = np.zeros_like(quantity)

    breakeven_fig = go.Figure()
    breakeven_fig.add_trace(go.Scatter(x=quantity, y=total_revenue, mode='lines', name='הכנסות', line=dict(color='#007bff')))
    breakeven_fig.add_trace(go.Scatter(x=quantity, y=total_costs, mode='lines', name='הוצאות', line=dict(color='#dc3545')))
    breakeven_fig.add_vline(x=breakeven_total_tickets, line_dash="dash", line_color="green",
                            annotation_text=f"נקודת איזון: {locale.format_string('%.0f', breakeven_total_tickets, grouping=True)} כרטיסים לשנה",
                            annotation_position="top right")
    breakeven_fig.update_layout(title='גרף נקודת איזון', xaxis_title='מספר כרטיסים לשנה', yaxis_title='ש״ח',
                                font=dict(size=14))

    # גרף זרם מזומנים (Cash Flow)
    cash_flow_fig = go.Figure()
    years = list(range(1, default_params['אורך מימון (שנים)'] + 1))
    cash_flow = [results['רווח לפני מס']] * default_params['אורך מימון (שנים)']
    cash_flow.insert(0, -results['עלויות הקמה'])  # השקעה התחלתית
    cumulative_cash_flow = np.cumsum(cash_flow)

    cash_flow_fig.add_trace(go.Bar(x=["השקעה"] + [f"שנה {year}" for year in years], y=cash_flow, name='תזרים מזומנים', marker=dict(color='#28a745')))
    cash_flow_fig.add_trace(
        go.Scatter(x=["השקעה"] + [f"שנה {year}" for year in years], y=cumulative_cash_flow, mode='lines+markers',
                   name='תזרים מזומנים מצטבר', line=dict(color='#007bff')))
    cash_flow_fig.update_layout(title='תזרים מזומנים', xaxis_title='שנים', yaxis_title='ש״ח', font=dict(size=14))

    # גרף שיעור רווחיות (Profit Margin)
    profit_margin = (results['רווח לפני מס'] / results['הכנסות שנתיות']) * 100 if results['הכנסות שנתיות'] != 0 else 0
    profit_margin_fig = go.Figure()
    profit_margin_fig.add_trace(go.Indicator(
        mode="gauge+number",
        value=profit_margin,
        title={'text': "שיעור רווחיות (%)"},
        gauge={
            'axis': {'range': [0, 100]},
            'steps': [
                {'range': [0, 20], 'color': "red"},
                {'range': [20, 60], 'color': "yellow"},
                {'range': [60, 100], 'color': "green"}
            ],
            'threshold': {'line': {'color': "black", 'width': 4}, 'thickness': 0.75, 'value': profit_margin}
        }
    ))
    profit_margin_fig.update_layout(font=dict(size=14))

    # גרף ניתוח רגישות
    sensitivity_fig = st.plotly_chart(update_sensitivity_graph('מספר מבקרים ביום רגיל', {'params': default_params}), use_container_width=True)

    sensitivity_param_dropdown = st.selectbox("בחר פרמטר לניתוח רגישות:", options=list(default_params.keys()))


    sensitivity_layout = st.container()
    with sensitivity_layout:
        sensitivity_param_dropdown
        sensitivity_fig

    # ניתוח רגישות מתקדם (הצלבה בין פרמטרים)
    advanced_sensitivity_layout = render_advanced_sensitivity_layout()

    # הצגת הגרפים
    st.plotly_chart(waterfall_fig, use_container_width=True)
    st.plotly_chart(income_pie_fig, use_container_width=True)
    st.plotly_chart(expense_pie_fig, use_container_width=True)
    st.plotly_chart(breakeven_fig, use_container_width=True)
    st.plotly_chart(cash_flow_fig, use_container_width=True)
    st.plotly_chart(profit_margin_fig, use_container_width=True)
    sensitivity_layout
    advanced_sensitivity_layout


# פונקציה לניתוח רגישות מתקדם (הצלבה בין פרמטרים)
def perform_advanced_sensitivity_analysis(current_params, param1, param2, range1, range2):
    profits = []
    combinations = []
    for val1 in range1:
        for val2 in range2:
            temp_params = current_params.copy()
            temp_params[param1] = val1
            temp_params[param2] = val2
            try:
                res = calculate_results(temp_params)
                profits.append(res['רווח לפני מס'])
                combinations.append((val1, val2))
            except KeyError as e:
                # אם יש פרמטר חסר, נמשיך בלי להוסיף
                print(f"Missing key: {e}")
                continue
    return combinations, profits

# פונקציה להצגת ניתוח רגישות מתקדם
def render_advanced_sensitivity_layout():
    sensitivity_params = [key for key in default_params.keys() if isinstance(default_params[key], (int, float))]

    param1 = st.selectbox("בחר פרמטר 1 לניתוח רגישות:", options=sensitivity_params)
    param2 = st.selectbox("בחר פרמטר 2 לניתוח רגישות:", options=sensitivity_params)

    advanced_sensitivity_fig = st.plotly_chart(update_advanced_sensitivity_graph(param1, param2, {'params': default_params}), use_container_width=True)

    advanced_sensitivity_info = st.container()
    with advanced_sensitivity_info:
        st.write("")  # Placeholder for sensitivity info


    advanced_sensitivity_layout = st.container()
    with advanced_sensitivity_layout:
        param1
        param2
        advanced_sensitivity_fig
        advanced_sensitivity_info

    return advanced_sensitivity_layout


# פונקציה לייצוא תוצאות לאקסל
def generate_excel(results):
    if not results:
        return None

    # יצירת DataFrame לתוצאות כלליות
    general_data = pd.DataFrame([
        {'מדד': key, 'ערך': value}
        for key, value in results.items() if not isinstance(value, dict)
    ])

    # יצירת DataFrame להכנסות לפי קטגוריות
    income_df = pd.DataFrame(list(results['הכנסות לפי קטגוריות'].items()), columns=['קטגוריה', 'סכום'])

    # יצירת DataFrame להוצאות לפי קטגוריות
    expense_df = pd.DataFrame(list(results['הוצאות לפי קטגוריות'].items()), columns=['קטגוריה', 'סכום'])

    # כתיבת הנתונים לאקסל
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        general_data.to_excel(writer, sheet_name='תוצאות כלליות', index=False)
        income_df.to_excel(writer, sheet_name='הכנסות לפי קטגוריות', index=False)
        expense_df.to_excel(writer, sheet_name='הוצאות לפי קטגוריות', index=False)
    output.seek(0)
    return output

# קולבק לניתוח רגישות חד-פרמטרי
def update_sensitivity_graph(param, store_data):
    if not store_data or 'params' not in store_data:
        return {}
    params = store_data['params']
    if param not in params:
        return {}
    # טווח שינוי של +/- 20%
    base_value = params.get(param, default_params.get(param, 1))
    if base_value == 0:
        param_values = np.array([0])
    else:
        param_values = np.linspace(base_value * 0.8, base_value * 1.2, 10)
    profits = []
    for val in param_values:
        temp_params = params.copy()
        temp_params[param] = val
        try:
            res = calculate_results(temp_params)
            profits.append(res['רווח לפני מס'])
        except KeyError as e:
            # אם יש פרמטר חסר, נמשיך בלי להוסיף
            print(f"Missing key: {e}")
            profits.append(0)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=param_values, y=profits, mode='lines+markers', name='רווח לפני מס'))
    fig.update_layout(
        title=f'ניתוח רגישות - {param}',
        xaxis_title=f"{param} (ש״ח)" if 'עלות' in param or 'מחיר' in param else param,
        yaxis_title='רווח לפני מס (ש״ח)',
        font=dict(size=14),
        hovermode="x unified"
    )
    return fig

# קולבק לניתוח רגישות מתקדם (הצלבה בין פרמטרים)
def update_advanced_sensitivity_graph(param1, param2, store_data):
    if not store_data or 'params' not in store_data:
        return {}
    params = store_data['params']
    if not param1 or not param2:
        return {}
    # טווח שינוי של +/- 20% לכל פרמטר
    base_value1 = params.get(param1, default_params.get(param1, 1))
    base_value2 = params.get(param2, default_params.get(param2, 1))
    if base_value1 == 0:
        param1_values = np.array([0])
    else:
        param1_values = np.linspace(base_value1 * 0.8, base_value1 * 1.2, 10)
    if base_value2 == 0:
        param2_values = np.array([0])
    else:
        param2_values = np.linspace(base_value2 * 0.8, base_value2 * 1.2, 10)
    combinations, profits = perform_advanced_sensitivity_analysis(params, param1, param2, param1_values, param2_values)

    if not combinations:
        return {}

    # יצירת DataFrame
    df = pd.DataFrame(combinations, columns=[param1, param2])
    df['רווח לפני מס'] = profits

    # יצירת heatmap
    fig = px.density_heatmap(
        df,
        x=param1,
        y=param2,
        z='רווח לפני מס',
        color_continuous_scale='Viridis',
        title=f'ניתוח רגישות מתקדם - {param1} מול {param2}',
        labels={param1: param1, param2: param2, 'רווח לפני מס': 'רווח לפני מס (ש״ח)'},
        nbinsx=10,
        nbinsy=10
    )
    fig.update_layout(font=dict(size=14))
    return fig

# קולבק להצגת מידע על ניתוח רגישות מתקדם
def display_sensitivity_info(clickData):
    if clickData is None:
        return ""
    param1_value = clickData['points'][0]['x']
    param2_value = clickData['points'][0]['y']
    profit = clickData['points'][0]['z']
    return f"רווח לפני מס: {locale.format_string('%.2f', profit, grouping=True)} ש״ח"

# פריסת האפליקציה (מותאם ל-Streamlit)
if 'results' not in st.session_state:
    render_parameters_tab()
else:
    results = st.session_state['results']
    tab = st.sidebar.radio("בחר טאב:", ("תוצאות", "גרפים"))
    if tab == "תוצאות":
        render_results_tab(results)

        # כפתור לייצוא לאקסל
        if st.button("ייצא לאקסל"):
            excel_io = generate_excel(results)
            st.download_button(
                label="הורדת קובץ אקסל",
                data=excel_io,
                file_name="gymboree_business_plan_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
       elif tab == "גרפים":
        render_charts_tab(results)

# הצגת טאב הגרפים (מותאם ל-Streamlit)
def render_charts_tab(results):
    """
    מציג את טאב הגרפים.

    Args:
        results (dict): תוצאות החישוב.
    """

    if not results:
        st.write("אנא הזן את הפרמטרים ולחץ על 'חשב' כדי לראות את הגרפים.")
        return

    # גרף מפל
    waterfall_fig = go.Figure(go.Waterfall(
        name="",
        orientation="v",
        measure=["relative", "relative", "relative", "relative", "total"],
        x=["הכנסות", "הוצאות משתנות", "הוצאות קבועות", "תשלומי הלוואה", "רווח לפני מס"],
        textposition="outside",
        text=[
            f"{locale.format_string('%.0f', results['הכנסות שנתיות'], grouping=True)} ש״ח",
            f"-{locale.format_string('%.0f', results['הוצאות משתנות'], grouping=True)} ש״ח",
            f"-{locale.format_string('%.0f', results['הוצאות קבועות'], grouping=True)} ש״ח",
            f"-{locale.format_string('%.0f', results['תשלומי הלוואה'], grouping=True)} ש״ח",
            f"{locale.format_string('%.0f', results['רווח לפני מס'], grouping=True)} ש״ח"
        ],
        y=[
            results['הכנסות שנתיות'],
            -results['הוצאות משתנות'],
            -results['הוצאות קבועות'],
            -results['תשלומי הלוואה'],
            results['רווח לפני מס']
        ],
        connector={"line": {"color": "rgb(63, 63, 63)"}},
        increasing={"marker": {"color": "#007bff"}},
        decreasing={"marker": {"color": "#dc3545"}},
        totals={"marker": {"color": "#28a745"}},
    ))
    waterfall_fig.update_layout(
        title="מפל הכנסות והוצאות",
        xaxis_title="",
        yaxis_title="ש״ח",
        font=dict(size=14),
        hovermode="x unified"
    )

    # גרף הכנסות לפי קטגוריות (גרף עוגה)
    income_df = pd.DataFrame(list(results['הכנסות לפי קטגוריות'].items()), columns=['קטגוריה', 'סכום'])
    income_pie_fig = px.pie(income_df, values='סכום', names='קטגוריה', title='התפלגות ההכנסות')
    income_pie_fig.update_traces(marker=dict(colors=['#007bff', '#28a745', '#dc3545', '#28a745', '#dc3545']))
    income_pie_fig.update_layout(font=dict(size=14))

    # גרף הוצאות לפי קטגוריות (גרף עוגה)
    expense_df = pd.DataFrame(list(results['הוצאות לפי קטגוריות'].items()), columns=['קטגוריה', 'סכום'])
    expense_pie_fig = px.pie(expense_df, values='סכום', names='קטגוריה', title='התפלגות ההוצאות')
    expense_pie_fig.update_layout(font=dict(size=14))

    # גרף נקודת איזון
    breakeven_total_tickets = results['נקודת איזון (מספר כרטיסים לשנה)']
    average_ticket_price = results['הכנסות שנתיות'] / (
            default_params['מספר מבקרים ביום רגיל'] * 191 + default_params['מספר מבקרים ביום חופשה/חג'] * 100) if (
            default_params['מספר מבקרים ביום רגיל'] * 191 + default_params['מספר מבקרים ביום חופשה/חג'] * 100) != 0 else 0

    quantity = np.linspace(0, breakeven_total_tickets * 1.5, 100)
    total_revenue = average_ticket_price * quantity
    if breakeven_total_tickets != 0:
        total_costs = (results['הוצאות משתנות'] + results['הוצאות קבועות'] + results['תשלומי הלוואה']) * quantity / breakeven_total_tickets
    else:
        total_costs = np.zeros_like(quantity)

    breakeven_fig = go.Figure()
    breakeven_fig.add_trace(go.Scatter(x=quantity, y=total_revenue, mode='lines', name='הכנסות', line=dict(color='#007bff')))
    breakeven_fig.add_trace(go.Scatter(x=quantity, y=total_costs, mode='lines', name='הוצאות', line=dict(color='#dc3545')))
    breakeven_fig.add_vline(x=breakeven_total_tickets, line_dash="dash", line_color="green",
                            annotation_text=f"נקודת איזון: {locale.format_string('%.0f', breakeven_total_tickets, grouping=True)} כרטיסים לשנה",
                            annotation_position="top right")
    breakeven_fig.update_layout(title='גרף נקודת איזון', xaxis_title='מספר כרטיסים לשנה', yaxis_title='ש״ח',
                                font=dict(size=14))

    # גרף זרם מזומנים (Cash Flow)
    cash_flow_fig = go.Figure()
    years = list(range(1, default_params['אורך מימון (שנים)'] + 1))
    cash_flow = [results['רווח לפני מס']] * default_params['אורך מימון (שנים)']
    cash_flow.insert(0, -results['עלויות הקמה'])  # השקעה התחלתית
    cumulative_cash_flow = np.cumsum(cash_flow)

    cash_flow_fig.add_trace(go.Bar(x=["השקעה"] + [f"שנה {year}" for year in years], y=cash_flow, name='תזרים מזומנים', marker=dict(color='#28a745')))
    cash_flow_fig.add_trace(
        go.Scatter(x=["השקעה"] + [f"שנה {year}" for year in years], y=cumulative_cash_flow, mode='lines+markers',
                   name='תזרים מזומנים מצטבר', line=dict(color='#007bff')))
    cash_flow_fig.update_layout(title='תזרים מזומנים', xaxis_title='שנים', yaxis_title='ש״ח', font=dict(size=14))

    # גרף שיעור רווחיות (Profit Margin)
    profit_margin = (results['רווח לפני מס'] / results['הכנסות שנתיות']) * 100 if results['הכנסות שנתיות'] != 0 else 0
    profit_margin_fig = go.Figure()
    profit_margin_fig.add_trace(go.Indicator(
        mode="gauge+number",
        value=profit_margin,
        title={'text': "שיעור רווחיות (%)"},
        gauge={
            'axis': {'range': [0, 100]},
            'steps': [
                {'range': [0, 20], 'color': "red"},
                {'range': [20, 60], 'color': "yellow"},
                {'range': [60, 100], 'color': "green"}
            ],
            'threshold': {'line': {'color': "black", 'width': 4}, 'thickness': 0.75, 'value': profit_margin}
        }
    ))
    profit_margin_fig.update_layout(font=dict(size=14))

    # גרף ניתוח רגישות
    sensitivity_param_dropdown = st.selectbox("בחר פרמטר לניתוח רגישות:", options=list(default_params.keys()))
    sensitivity_fig = st.plotly_chart(update_sensitivity_graph(sensitivity_param_dropdown, {'params': default_params}), use_container_width=True)

    # ניתוח רגישות מתקדם (הצלבה בין פרמטרים)
    advanced_sensitivity_layout = render_advanced_sensitivity_layout()

    # הצגת הגרפים
    st.plotly_chart(waterfall_fig, use_container_width=True)
    st.plotly_chart(income_pie_fig, use_container_width=True)
    st.plotly_chart(expense_pie_fig, use_container_width=True)
    st.plotly_chart(breakeven_fig, use_container_width=True)
    st.plotly_chart(cash_flow_fig, use_container_width=True)
    st.plotly_chart(profit_margin_fig, use_container_width=True)
    sensitivity_fig # Removed layout
    advanced_sensitivity_layout


# פונקציה לניתוח רגישות מתקדם (הצלבה בין פרמטרים)
def perform_advanced_sensitivity_analysis(current_params, param1, param2, range1, range2):
    profits = []
    combinations = []
    for val1 in range1:
        for val2 in range2:
            temp_params = current_params.copy()
            temp_params[param1] = val1
            temp_params[param2] = val2
            try:
                res = calculate_results(temp_params)
                profits.append(res['רווח לפני מס'])
                combinations.append((val1, val2))
            except KeyError as e:
                # אם יש פרמטר חסר, נמשיך בלי להוסיף
                print(f"Missing key: {e}")
                continue
    return combinations, profits

# פונקציה להצגת ניתוח רגישות מתקדם
def render_advanced_sensitivity_layout():
    sensitivity_params = [key for key in default_params.keys() if isinstance(default_params[key], (int, float))]

    param1 = st.selectbox("בחר פרמטר 1 לניתוח רגישות:", options=sensitivity_params)
    param2 = st.selectbox("בחר פרמטר 2 לניתוח רגישות:", options=sensitivity_params)

    advanced_sensitivity_fig = st.plotly_chart(update_advanced_sensitivity_graph(param1, param2, {'params': default_params}), use_container_width=True)

    advanced_sensitivity_info = st.container()
    with advanced_sensitivity_info:
        st.write("")  # Placeholder for sensitivity info


    advanced_sensitivity_layout = st.container()
    with advanced_sensitivity_layout:
        param1
        param2
        advanced_sensitivity_fig
        advanced_sensitivity_info

    return advanced_sensitivity_layout


# פונקציה לייצוא תוצאות לאקסל
def generate_excel(results):
    if not results:
        return None

    # יצירת DataFrame לתוצאות כלליות
    general_data = pd.DataFrame([
        {'מדד': key, 'ערך': value}
        for key, value in results.items() if not isinstance(value, dict)
    ])

    # יצירת DataFrame להכנסות לפי קטגוריות
    income_df = pd.DataFrame(list(results['הכנסות לפי קטגוריות'].items()), columns=['קטגוריה', 'סכום'])

    # יצירת DataFrame להוצאות לפי קטגוריות
    expense_df = pd.DataFrame(list(results['הוצאות לפי קטגוריות'].items()), columns=['קטגוריה', 'סכום'])

    # כתיבת הנתונים לאקסל
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        general_data.to_excel(writer, sheet_name='תוצאות כלליות', index=False)
        income_df.to_excel(writer, sheet_name='הכנסות לפי קטגוריות', index=False)
        expense_df.to_excel(writer, sheet_name='הוצאות לפי קטגוריות', index=False)
    output.seek(0)
    return output

# קולבק לניתוח רגישות חד-פרמטרי
def update_sensitivity_graph(param, store_data):
    if not store_data or 'params' not in store_data:
        return {}
    params = store_data['params']
    if param not in params:
        return {}
    # טווח שינוי של +/- 20%
    base_value = params.get(param, default_params.get(param, 1))
    if base_value == 0:
        param_values = np.array([0])
    else:
        param_values = np.linspace(base_value * 0.8, base_value * 1.2, 10)
    profits = []
    for val in param_values:
        temp_params = params.copy()
        temp_params[param] = val
        try:
            res = calculate_results(temp_params)
            profits.append(res['רווח לפני מס'])
        except KeyError as e:
            # אם יש פרמטר חסר, נמשיך בלי להוסיף
            print(f"Missing key: {e}")
            profits.append(0)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=param_values, y=profits, mode='lines+markers', name='רווח לפני מס'))
    fig.update_layout(
        title=f'ניתוח רגישות - {param}',
        xaxis_title=f"{param} (ש״ח)" if 'עלות' in param or 'מחיר' in param else param,
        yaxis_title='רווח לפני מס (ש״ח)',
        font=dict(size=14),
        hovermode="x unified"
    )
    return fig

# קולבק לניתוח רגישות מתקדם (הצלבה בין פרמטרים)
def update_advanced_sensitivity_graph(param1, param2, store_data):
    if not store_data or 'params' not in store_data:
        return {}
    params = store_data['params']
    if not param1 or not param2:
        return {}
    # טווח שינוי של +/- 20% לכל פרמטר
    base_value1 = params.get(param1, default_params.get(param1, 1))
    base_value2 = params.get(param2, default_params.get(param2, 1))
    if base_value1 == 0:
        param1_values = np.array([0])
    else:
        param1_values = np.linspace(base_value1 * 0.8, base_value1 * 1.2, 10)
    if base_value2 == 0:
        param2_values = np.array([0])
    else:
        param2_values = np.linspace(base_value2 * 0.8, base_value2 * 1.2, 10)
    combinations, profits = perform_advanced_sensitivity_analysis(params, param1, param2, param1_values, param2_values)

    if not combinations:
        return {}

    # יצירת DataFrame
    df = pd.DataFrame(combinations, columns=[param1, param2])
    df['רווח לפני מס'] = profits

    # יצירת heatmap
    fig = px.density_heatmap(
        df,
        x=param1,
        y=param2,
        z='רווח לפני מס',
        color_continuous_scale='Viridis',
        title=f'ניתוח רגישות מתקדם - {param1} מול {param2}',
        labels={param1: param1, param2: param2, 'רווח לפני מס': 'רווח לפני מס (ש״ח)'},
        nbinsx=10,
        nbinsy=10
    )
    fig.update_layout(font=dict(size=14))
    return fig

# קולבק להצגת מידע על ניתוח רגישות מתקדם
def display_sensitivity_info(clickData):
    if clickData is None:
        return ""
    param1_value = clickData['points'][0]['x']
    param2_value = clickData['points'][0]['y']
    profit = clickData['points'][0]['z']
    return f"רווח לפני מס: {locale.format_string('%.2f', profit, grouping=True)} ש״ח"

# פריסת האפליקציה (מותאם ל-Streamlit)
if 'results' not in st.session_state:
    render_parameters_tab()
else:
    results = st.session_state['results']
    tab = st.sidebar.radio("בחר טאב:", ("תוצאות", "גרפים"))
    if tab == "תוצאות":
        render_results_tab(results)

        # כפתור לייצוא לאקסל
        if st.button("ייצא לאקסל"):
            excel_io = generate_excel(results)
            st.download_button(
                label="הורדת קובץ אקסל",
                data=excel_io,
                file_name="gymboree_business_plan_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    elif tab == "גרפים":
        render_charts_tab(results)
