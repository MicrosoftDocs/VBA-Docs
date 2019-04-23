---
title: Worksheet.StandardHeight property (Excel)
keywords: vbaxl10.chm175129
f1_keywords:
- vbaxl10.chm175129
ms.prod: excel
api_name:
- Excel.Worksheet.StandardHeight
ms.assetid: 43dd7321-5df7-2cda-9e51-75f81ab203f2
ms.date: 06/08/2017
localization_priority: Normal
---


# Worksheet.StandardHeight property (Excel)

Returns the standard (default) height of all the rows in the worksheet, in points. Read-only  **Double**.


## Syntax

_expression_. `StandardHeight`

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example sets the height of row one on Sheet1 to the standard height.


```vb
Worksheets("Sheet1").Rows(1).RowHeight = _ 
 Worksheets("Sheet1").StandardHeight
```


## See also


[Worksheet Object](Excel.Worksheet.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]