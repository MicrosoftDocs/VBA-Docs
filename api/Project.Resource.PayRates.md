---
title: Resource.PayRates property (Project)
ms.prod: project-server
api_name:
- Project.Resource.PayRates
ms.assetid: bd01dd18-bbf4-52d5-bc37-d525603fcb8e
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.PayRates property (Project)

Gets a **[PayRates](Project.PayRate.md)** collection that represents the various pay rates on the cost rate table for a resource. Read-only **PayRates**.


## Syntax

_expression_. `PayRates`

_expression_ A variable that represents a [Resource](./Project.Resource.md) object.


## Remarks

For the **Resource** object, the **PayRates** property returns pay rates for cost rate table A, the default table.


## Example

The following example lists the standard pay rates for all the cost rate tables of the resource in the active cell.


```vb
Sub ListPayRates() 
 Dim CRT As CostRateTable, PR As PayRate 
 Dim Rates As String 
 
 For Each CRT In ActiveCell.Resource.CostRateTables 
 For Each PR In CRT.PayRates 
 Rates = Rates & "CostRateTable " & CRT.Name & ": " & _ 
 PR.StandardRate & " (Effective " & PR.EffectiveDate & _ 
 ")" & vbCrLf 
 Next PR 
 Next CRT 
 
 MsgBox Rates 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]