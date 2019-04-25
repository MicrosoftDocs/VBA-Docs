---
title: Font.Italic property (Excel)
keywords: vbaxl10.chm559078
f1_keywords:
- vbaxl10.chm559078
ms.prod: excel
api_name:
- Excel.Font.Italic
ms.assetid: 9d249157-9c8a-79ec-9b70-021c19ea1336
ms.date: 04/26/2019
localization_priority: Normal
---


# Font.Italic property (Excel)

**True** if the font style is italic. Read/write **Boolean**.


## Syntax

_expression_.**Italic**

_expression_ A variable that represents a **[Font](excel.font(object).md)** object.


## Example

This example sets the font style to italic for the range A1:A5 on Sheet1.

```vb
Worksheets("Sheet1").Range("A1:A5").Font.Italic = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
