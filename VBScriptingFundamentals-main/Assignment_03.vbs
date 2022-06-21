option explicit

Const SITE_TITLE = "www.code.com"
Const TAXES = 0.2, RETIRE = 0.05, HEALTH = 200, SALARY = 60000

Dim Monthly
Dim Total_Monthly_Deductables, Monthly_Wage, Total_After_Deductables

Monthly_Wage = SALARY / 12
Total_Monthly_Deductables = (Monthly_Wage * TAXES) + (Monthly_Wage * RETIRE) + HEALTH
Total_After_Deductables = Monthly_Wage - Total_Monthly_Deductables


MsgBox "Adam receives $" & Total_After_Deductables & " as monthly salary after all deductions!", 0, SITE_TITLE
