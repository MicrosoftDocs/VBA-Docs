---
title: DefaultWebOptions.Fonts property (Excel)
keywords: vbaxl10.chm660088
f1_keywords:
- vbaxl10.chm660088
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.Fonts
ms.assetid: a1b79e75-98a4-a784-522c-0aa72fd65b5c
ms.date: 06/08/2017
localization_priority: Normal
---


# DefaultWebOptions.Fonts property (Excel)

Returns the  **[WebPageFonts](Office.WebPageFonts.md)** collection representing the set of fonts Microsoft Excel uses when you open a Web page in Excel and there is either no font information specified in the Web page, or the current default font can't display the character set in the Web page. Read-only.


## Syntax

_expression_. `Fonts`

_expression_ A variable that represents a [DefaultWebOptions](Excel.DefaultWebOptions.md) object.


## Example

This example sets the default fixed-width font for the English/Western European/Other Latin Script character set to Courier New, 14 points.


```vb
With Application.DefaultWebOptions _ 
    .Fonts(msoCharacterSetEnglishWesternEuropeanOtherLatinScript) 
        .FixedWidthFont = "Courier New" 
        .FixedWidthFontSize = 14 
End With
```


## See also


[DefaultWebOptions Object](Excel.DefaultWebOptions.md)

