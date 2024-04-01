---
title: Font.Subscript property (Excel)
keywords: vbaxl10.chm559084
f1_keywords:
- vbaxl10.chm559084
api_name:
- Excel.Font.Subscript
ms.assetid: fb98ecb9-9653-4b5e-f3e1-838309069810
ms.date: 04/26/2019
ms.localizationpriority: medium
---


# Font.Subscript property (Excel)

**True** if the font is formatted as subscript. **False** by default. Read/write **Variant**.


## Syntax

_expression_.**Subscript**

_expression_ A variable that represents a **[Font](excel.font(object).md)** object.


## Example

This example makes the second character in cell A1 a subscript character.

```vb
Worksheets("Sheet1").Range("A1") _ 
 .Characters(2, 1).Font.Subscript = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]