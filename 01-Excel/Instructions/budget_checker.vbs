Sub BudgetChecker()
    Dim budget as Double
    Dim price as Double
    dim fee as Double

    budget = Cells(3,3).Value
    price = Cells(3,6).Value
    fee = Cells(3,8).Value

    Dim total as Double
    total = price*(1+fee)
    Range("L3").Value = total