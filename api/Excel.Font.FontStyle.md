---
title: Font.FontStyle property (Excel)
keywords: vbaxl10.chm559077
f1_keywords:
- vbaxl10.chm559077
ms.prod: excel
api_name:
- Excel.Font.FontStyle
ms.assetid: 17e5989e-09a5-dabb-4989-82daf3aa0295
ms.date: 04/26/2019
localization_priority: Normal
---


# Font.FontStyle property (Excel)

Returns or sets the font style. Read/write **String**.


## Syntax

_expression_.**FontStyle**

_expression_ A variable that represents a **[Font](excel.font(object).md)** object.


## Remarks

Changing this property may affect other **Font** properties (such as **[Bold](Excel.Font.Bold.md)** and **[Italic](Excel.Font.Italic.md)**). Acceptable values are Regular, Italic, Bold, and Bold Italic.


## Example

This example sets the font style for cell A1 on Sheet1 to bold and italic.

```vb
Worksheets("Sheet1").Range("A1").Font.FontStyle = "Bold Italic"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
