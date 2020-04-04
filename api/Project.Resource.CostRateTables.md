---
title: Resource.CostRateTables property (Project)
ms.prod: project-server
api_name:
- Project.Resource.CostRateTables
ms.assetid: 604d89ee-a16e-812e-0459-b93ed096340e
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.CostRateTables property (Project)

Gets a **[CostRateTables](Project.CostRateTable.md)** collection representing the cost rate tables for the resource. Read-only **CostRateTables**.


## Syntax

_expression_. `CostRateTables`

_expression_ A variable that represents a [Resource](./Project.Resource.md) object.


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