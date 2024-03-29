---
title: Range.Justify method (Excel)
keywords: vbaxl10.chm144152
f1_keywords:
- vbaxl10.chm144152
api_name:
- Excel.Range.Justify
ms.assetid: f8b4d48b-8cbb-977a-fd44-d354661182d2
ms.date: 05/11/2019
ms.localizationpriority: medium
---


# Range.Justify method (Excel)

Rearranges the text in a range so that it fills the range evenly.


## Syntax

_expression_.**Justify**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Return value

Variant


## Remarks

If the range isn't large enough, Microsoft Excel displays a message telling you that text will extend below the range. If you choose the **OK** button, justified text replaces the contents in cells that extend beyond the selected range. To prevent this message from appearing, set the **[DisplayAlerts](Excel.Application.DisplayAlerts.md)** property to **False**. After you set this property, text will always replace the contents in cells below the range.


## Example

This example justifies the text in cell A1 on Sheet1.

```vb
Worksheets("Sheet1").Range("A1").Justify
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]