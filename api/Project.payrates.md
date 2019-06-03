---
title: PayRates object (Project)
ms.prod: project-server
ms.assetid: 7aa54cc3-4e39-e3b1-f3aa-7599ac88d22a
ms.date: 06/08/2017
localization_priority: Normal
---


# PayRates object (Project)

Contains a collection of  **[PayRate](Project.PayRate.md)** objects.
 


## Example

 **Using the PayRate Object**
 

 
Use  **PayRates** (*Index* ), where*Index* is the pay rate index number or date for which to return the rates in effect, to return a single **PayRate** object. The following example returns the standard pay rate for Tamara's first row of rates in cost rate table C.
 

 



```vb
ActiveProject.Resources("Tamara").CostRateTables("C").PayRates(1).StandardRate
```

 **Using the PayRates Collection**
 

 
Use the  **[PayRates](Project.CostRateTable.PayRates.md)** property to return a **PayRates** collection. The following example lists the standard pay rates for all the cost rate tables of the resource in the active cell.
 

 



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

Use the  **[Add](Project.PayRates.Add.md)** method to add a **PayRate** object to the **PayRates** collection. The following example adds a line to Tamara's cost rate table "C" with an effective date of September 1, 2012, a standard rate of $40.00 per hour, an overtime rate of $60.00 per hour, and a per-use cost of $0.
 

 



```vb
ActiveProject.Resources("Tamara").CostRateTables("C").PayRates.Add "9/1/2012", "$40/h", "$60/h", "$0"
```


## Methods



|Name|
|:-----|
|[Add](Project.PayRates.Add.md)|

## Properties



|Name|
|:-----|
|[Application](Project.PayRates.Application.md)|
|[Count](Project.PayRates.Count.md)|
|[Item](Project.PayRates.Item.md)|
|[Parent](Project.PayRates.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]