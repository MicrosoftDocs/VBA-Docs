---
title: Font.Underline property (Excel)
keywords: vbaxl10.chm559086
f1_keywords:
- vbaxl10.chm559086
ms.prod: excel
api_name:
- Excel.Font.Underline
ms.assetid: 81a2bdd2-bebd-b3ca-e0c3-6dd55280fcc0
ms.date: 04/26/2019
localization_priority: Normal
---


# Font.Underline property (Excel)

Returns or sets the type of underline applied to the font. Read/write **Variant**.


## Syntax

_expression_.**Underline**

_expression_ A variable that represents a **[Font](excel.font(object).md)** object.


## Remarks

Can be one of the **[XlUnderlineStyle](Excel.XlUnderlineStyle.md)** constants.

## Example

This example sets the font in the active cell on Sheet1 to single underline.

```vb
Worksheets("Sheet1").Activate 
ActiveCell.Font.Underline = xlUnderlineStyleSingle
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
