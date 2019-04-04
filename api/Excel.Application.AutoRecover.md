---
title: Application.AutoRecover property (Excel)
keywords: vbaxl10.chm133276
f1_keywords:
- vbaxl10.chm133276
ms.prod: excel
api_name:
- Excel.Application.AutoRecover
ms.assetid: bc2453fa-4319-c1da-5ad5-2efb306c3063
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.AutoRecover property (Excel)

Returns an **[AutoRecover](Excel.AutoRecover.md)** object, which backs up all file formats on a timed interval.


## Syntax

_expression_.**AutoRecover**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

Valid time intervals are whole numbers from 1 to 120.


## Example

In this example, the **[Time](Excel.AutoRecover.Time.md)** property is used in conjunction with the **AutoRecover** property to set the time interval for Microsoft Excel to wait before saving another copy to five minutes.

```vb
Sub UseAutoRecover() 
 
 Application.AutoRecover.Time = 5 
 
 MsgBox "The time that will elapse between each automatic " & _ 
 "save has been set to " & _ 
 Application.AutoRecover.Time & " minutes." 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]