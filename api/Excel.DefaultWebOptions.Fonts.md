---
title: DefaultWebOptions.Fonts property (Excel)
keywords: vbaxl10.chm660088
f1_keywords:
- vbaxl10.chm660088
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.Fonts
ms.assetid: a1b79e75-98a4-a784-522c-0aa72fd65b5c
ms.date: 04/25/2019
localization_priority: Normal
---


# DefaultWebOptions.Fonts property (Excel)

Returns the **[WebPageFonts](Office.WebPageFonts.md)** collection representing the set of fonts Microsoft Excel uses when you open a webpage in Excel and there is either no font information specified on the webpage, or the current default font can't display the character set on the webpage. Read-only.


## Syntax

_expression_.**Fonts**

_expression_ A variable that represents a **[DefaultWebOptions](Excel.DefaultWebOptions.md)** object.


## Example

This example sets the default fixed-width font for the English/Western European/Other Latin Script character set to Courier New, 14 points.

```vb
With Application.DefaultWebOptions _ 
    .Fonts(msoCharacterSetEnglishWesternEuropeanOtherLatinScript) 
        .FixedWidthFont = "Courier New" 
        .FixedWidthFontSize = 14 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]