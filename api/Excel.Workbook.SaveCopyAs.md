---
title: Workbook.SaveCopyAs method (Excel)
keywords: vbaxl10.chm199146
f1_keywords:
- vbaxl10.chm199146
ms.prod: excel
api_name:
- Excel.Workbook.SaveCopyAs
ms.assetid: 84f58488-6a2b-7fef-1472-e1b9771a60b0
ms.date: 06/08/2017
localization_priority: Priority
---


# Workbook.SaveCopyAs method (Excel)

Saves a copy of the workbook to a file but doesn't modify the open workbook in memory.


## Syntax

_expression_. `SaveCopyAs`( `_Filename_` )

_expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **Variant**|Specifies the file name for the copy.|

## Example

This example saves a copy of the active workbook.


```vb
ActiveWorkbook.SaveCopyAs "C:\TEMP\XXXX.XLS"
```


## See also


[Workbook Object](Excel.Workbook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]