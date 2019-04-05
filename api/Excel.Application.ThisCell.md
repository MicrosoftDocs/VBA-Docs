---
title: Application.ThisCell property (Excel)
keywords: vbaxl10.chm133291
f1_keywords:
- vbaxl10.chm133291
ms.prod: excel
api_name:
- Excel.Application.ThisCell
ms.assetid: 83b9c009-7e01-4493-bda0-cd6246aba778
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.ThisCell property (Excel)

Returns the cell in which the user-defined function is being called from as a **[Range](Excel.Range(object).md)** object.


## Syntax

_expression_.**ThisCell**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

Users should not access properties or methods on the **Range** object when inside the user-defined function. Users can cache the **Range** object for later use and perform additional actions when the recalculation is finished.


## Example

In this example, a function called UseThisCell contains the **ThisCell** property to notify the user of the cell address.

```vb
Function UseThisCell() 
 MsgBox "The cell address is: " & _ 
 Application.ThisCell.Address 
End Function
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]