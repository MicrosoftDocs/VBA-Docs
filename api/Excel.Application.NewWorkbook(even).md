---
title: Application.NewWorkbook event (Excel)
keywords: vbaxl10.chm504073
f1_keywords:
- vbaxl10.chm504073
ms.prod: excel
api_name:
- Excel.Application.NewWorkbook
ms.assetid: a3c29269-af09-08da-f0c3-82e192aa896f
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.NewWorkbook event (Excel)

Occurs when a new workbook is created.


## Syntax

_expression_.**NewWorkbook** (_Wb_)

_expression_ An expression that returns an **[Application](Excel.Application(object).md)** object.


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



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]