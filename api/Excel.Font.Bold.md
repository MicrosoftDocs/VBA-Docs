---
title: Font.Bold property (Excel)
keywords: vbaxl10.chm559074
f1_keywords:
- vbaxl10.chm559074
ms.prod: excel
api_name:
- Excel.Font.Bold
ms.assetid: 7343989f-f973-0b1d-e595-c625ef2e0c15
ms.date: 04/26/2019
localization_priority: Normal
---


# Font.Bold property (Excel)

**True** if the font is bold. Read/write **Variant**.


## Syntax

_expression_.**Bold**

_expression_ A variable that represents a **[Font](excel.font(object).md)** object.


## Example

This example sets the font to bold for the range A1:A5 on Sheet1.

```vb
Worksheets("Sheet1").Range("A1:A5").Font.Bold = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
