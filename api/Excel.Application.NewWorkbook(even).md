---
title: Application.NewWorkbook Event (Excel)
keywords: vbaxl10.chm504073
f1_keywords:
- vbaxl10.chm504073
ms.prod: excel
api_name:
- Excel.Application.NewWorkbook
ms.assetid: a3c29269-af09-08da-f0c3-82e192aa896f
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.NewWorkbook Event (Excel)

Occurs when a new workbook is created.


## Syntax

_expression_. `NewWorkbook`( `_Wb_` )

 _expression_ An expression that returns a '[Application](Excel.Application(object).md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The new workbook.|

## Example

This example arranges open windows when a new workbook is created.


```vb
Private Sub App_NewWorkbook(ByVal Wb As Workbook) 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]