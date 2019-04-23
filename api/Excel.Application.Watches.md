---
title: Application.Watches property (Excel)
keywords: vbaxl10.chm133267
f1_keywords:
- vbaxl10.chm133267
ms.prod: excel
api_name:
- Excel.Application.Watches
ms.assetid: 487c5cad-67bf-3bc9-dbc4-6bd8a105ed5e
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.Watches property (Excel)

Returns a **[Watches](Excel.Watches.md)** object representing a range that is tracked when the worksheet is recalculated.


## Syntax

_expression_.**Watches**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example creates a summation formula in cell A3, and then adds this cell to the Watch window.

```vb
Sub AddWatch() 
 With Application 
 .Range("A1").Formula = 1 
 .Range("A2").Formula = 2 
 .Range("A3").Formula = "=Sum(A1:A2)" 
 .Range("A3").Select 
 .Watches.Add Source:=ActiveCell 
 End With 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]