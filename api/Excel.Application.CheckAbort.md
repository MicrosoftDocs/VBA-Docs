---
title: Application.CheckAbort method (Excel)
keywords: vbaxl10.chm133279
f1_keywords:
- vbaxl10.chm133279
ms.prod: excel
api_name:
- Excel.Application.CheckAbort
ms.assetid: e407aeff-b401-029a-9ada-8f11eef54fb0
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.CheckAbort method (Excel)

Stops recalculation in a Microsoft Excel application.


## Syntax

_expression_.**CheckAbort** (_KeepAbort_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeepAbort_|Optional| **Boolean**|Allows recalculation to be performed for a range.|

## Example

In this example, Excel stops recalculation in the application, except for cell A10. For you to be able to see the results of this example, other calculations should exist in the application that will allow you to see the differences between the cell designated to continue recalculating and other cells.

```vb
Sub UseCheckAbort() 
 
 Dim rngSubtotal As Variant 
 Set rngSubtotal = Application.Range("A10") 
 
 ' Stop recalculation except for designated cell. 
 Application.CheckAbort KeepAbort:=rngSubtotal 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]