---
title: Workbook.InactiveListBorderVisible property (Excel)
keywords: vbaxl10.chm199229
f1_keywords:
- vbaxl10.chm199229
ms.prod: excel
api_name:
- Excel.Workbook.InactiveListBorderVisible
ms.assetid: a6259862-9a29-f3a5-498f-633f51ec10e6
ms.date: 06/08/2017
localization_priority: Normal
---


# Workbook.InactiveListBorderVisible property (Excel)

A  **Boolean** value that specifies whether list borders are visible when a list is not active. Returns **True** if the border is visible. Read/write **Boolean**.


## Syntax

_expression_. `InactiveListBorderVisible`

_expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


## Remarks

Setting this property will affect all the lists that are on the worksheet.


## Example

The following example hides the borders of inactive lists in the workbook.


```vb
Sub HideListBorders() 
 
 ActiveWorkbook.InactiveListBorderVisible = False 
 
End Sub
```


## See also


[Workbook Object](Excel.Workbook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]