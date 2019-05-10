---
title: Range.Errors property (Excel)
keywords: vbaxl10.chm144235
f1_keywords:
- vbaxl10.chm144235
ms.prod: excel
api_name:
- Excel.Range.Errors
ms.assetid: 88dcc606-d412-a9ce-82bc-5fbba8baae87
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.Errors property (Excel)

Allows the user to access error checking options.


## Syntax

_expression_.**Errors**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

Reference the **[Errors](Excel.Errors.md)** object to view a list of index values associated with error checking options.


## Example

In this example, a number written as text is placed in cell A1. Microsoft Excel then determines if the number is written as text in cell A1 and notifies the user accordingly.

```vb
Sub CheckForErrors() 
 
 Range("A1").Formula = "'12" 
 
 If Range("A1").Errors.Item(xlNumberAsText).Value = True Then 
 MsgBox "The number is written as text." 
 Else 
 MsgBox "The number is not written as text." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]