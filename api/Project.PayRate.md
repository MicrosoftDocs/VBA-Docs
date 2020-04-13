---
title: PayRate object (Project)
ms.prod: project-server
api_name:
- Project.PayRate
ms.assetid: 4c8ba1f3-bf18-2179-5f50-c090c63e46b9
ms.date: 06/08/2017
localization_priority: Normal
---


# PayRate object (Project)


 

Represents a line of rates from the cost rate table of a resource. The **PayRate** object is a member of the **[PayRates](Project.payrates.md)** collection.
 
 **Using the PayRate Object**
 
Use  **PayRates** (*Index* ), where*Index* is the pay rate index number or date for which to return the rates in effect, to return a single **PayRate** object. The following example returns the standard pay rate for Tamara's first row of rates in cost rate table C.
 



```vb
ActiveProject.Resources("Tamara").CostRateTables("C").PayRates(1).StandardRate
```

 **Using the PayRates Collection**
 
Use the **[PayRates](Project.CostRateTable.PayRates.md)** property to return a **PayRates** collection. The following example lists the standard pay rates for all the cost rate tables of the resource in the active cell.
 



```vb
Dim CRT As CostRateTable
DIM PR As PayRate
Dim Rates As String

For Each CRT In ActiveCell.Resource.CostRateTables
    For Each PR In CRT.PayRates
        Rates = Rates & "CostRateTable " & CRT.Name & ": " & PR.StandardRate & vbCrLf
    Next PR
Next CRT
    
MsgBox Rates
```

Use the **[Add](Project.PayRates.Add.md)** method to add a **PayRate** object to the **PayRates** collection. The following example adds a line to Tamara's cost rate table "C" with an effective date of September 1, 2012, a standard rate of $40.00 per hour, an overtime rate of $60.00 per hour, and a per-use cost of $0.
 



```vb
ActiveProject.Resources("Tamara").CostRateTables("C").PayRates.Add "9/1/2012", "$40/h", "$60/h", "$0"
```


## Methods



|Name|
|:-----|
|[Delete](Project.PayRate.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](Project.PayRate.Application.md)|
|[CostPerUse](Project.PayRate.CostPerUse.md)|
|[EffectiveDate](Project.PayRate.EffectiveDate.md)|
|[Index](Project.PayRate.Index.md)|
|[OvertimeRate](Project.PayRate.OvertimeRate.md)|
|[Parent](Project.PayRate.Parent.md)|
|[StandardRate](Project.PayRate.StandardRate.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]