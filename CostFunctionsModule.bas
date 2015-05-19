Attribute VB_Name = "CostFunctionsModule"
' rev 5/19/2015 fixed 

Public Function foriegnExchange(USDollars As Double) As Double
    foriegnExchange = USDollars * 2.0175 + USDollars * 2.0175 * 0.0175 + USDollars * 0.025
End Function

Public Function CIF(USDollars As Double) As Double

    CIF = (USDollars + Freight(USDollars)) * 2.0175
End Function

Public Function Duty(USDollars As Double, dutyRate As Double) As Double
    Duty = (CIF(USDollars) * dutyRate)
End Function

Public Function ET(USDollars As Double) As Double
    ET = CIF(USDollars) * 0.02
End Function


Public Function GST(USDollars As Double, dutyRate As Double) As Double
    CIFAmount = CIF(USDollars)
    DutyAmount = Duty(USDollars, dutyRate)
    ETAmount = ET(USDollars)
    GST = (CIFAmount + DutyAmount + ETAmount) * 0.125
End Function

Public Function Freight(USDollars As Double) As Double
    Freight = USDollars * 0.05
End Function
Public Function CostBZD(USDollars As Double, dutyRate As Double) As Double
    CostBZD = foriegnExchange(USDollars) + Freight(USDollars) + Duty(USDollars, dutyRate) + ET(USDollars) + GST(USDollars, dutyRate)

End Function
