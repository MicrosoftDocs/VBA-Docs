---
title: FullSeriesCollection.Application property (Excel)
keywords: vbaxl10.chm943073
f1_keywords:
- vbaxl10.chm943073
ms.prod: excel
ms.assetid: 52dfb5aa-c6fb-201c-c1ed-880aff1efb45
ms.date: 06/08/2017
localization_priority: Normal
---


# FullSeriesCollection.Application property (Excel)

Returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[FullSeriesCollection](Excel.fullseriescollection.md)** object.


## Example

This example displays a message about the application that created  `myObject`.


```vb
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


## Property value

 **APPLICATION**



## See also



[FullSeriesCollection Object](Excel.fullseriescollection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]