option explicit

Const SITE_TITLE = "www.code.com"
Const TAXES = 0.2, RETIRE = 0.05, HEALTH = 200

Dim Monthly, Total_Monthly_Deductables, Monthly_Wage, Total_After_Deductables, Users_Name, Users_Salary

Users_Name = InputBox("Enter your name : ")
Users_Salary = InputBox("Enter your Salary: ")
Monthly_Wage = Users_Salary / 12
Total_Monthly_Deductables = (Monthly_Wage * TAXES) + (Monthly_Wage * RETIRE) + HEALTH
Total_After_Deductables = Monthly_Wage - Total_Monthly_Deductables


MsgBox Users_Name & " receives $" & Total_After_Deductables & " as monthly salary after all deductions!", 0, SITE_TITLE
