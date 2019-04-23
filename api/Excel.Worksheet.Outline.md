---
title: Worksheet.Outline property (Excel)
keywords: vbaxl10.chm175113
f1_keywords:
- vbaxl10.chm175113
ms.prod: excel
api_name:
- Excel.Worksheet.Outline
ms.assetid: e53d8038-f20b-9d55-1ee0-c5f6b4a099d4
ms.date: 06/08/2017
localization_priority: Normal
---


# Worksheet.Outline property (Excel)

Returns an  **[Outline](Excel.Outline.md)** object that represents the outline for the specified worksheet. Read-only.


## Syntax

_expression_. `Outline`

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example sets the outline on Sheet1 to use automatic styles.


```vb
Worksheets("Sheet1").Outline.AutomaticStyles = True
```


## See also


[Worksheet Object](Excel.Worksheet.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]