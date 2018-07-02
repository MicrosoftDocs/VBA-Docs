---
title: WebPageFont.ProportionalFont Property (Office)
keywords: vbaof11.chm224001
f1_keywords:
- vbaof11.chm224001
ms.prod: office
api_name:
- Office.WebPageFont.ProportionalFont
ms.assetid: fcefea5f-4c9f-c050-9599-fdf4c9269bdd
ms.date: 06/08/2017
---


# WebPageFont.ProportionalFont Property (Office)

Sets or gets the proportional font setting in the host application. Read/write.


## Syntax

 _expression_. `ProportionalFont`

 _expression_ A variable that represents a [WebPageFont](./Office.WebPageFont.md) object.


## Remarks

When you set the  **ProportionalFont** property, the host application does not check the value for validity.


## Example

This example sets the proportional font and proportional font size for the English/Western European/Other Latin Script character set in the active application.


```vb
Application.DefaultWebOptions. _ 
Fonts(msoCharacterSetEnglishWesternEuropeanOtherLatinScript) _ 
.ProportionalFont = "Tahoma" 
Application.DefaultWebOptions. _ 
Fonts(msoCharacterSetEnglishWesternEuropeanOtherLatinScript) _ 
.ProportionalFontSize = 14.5
```


## See also


[WebPageFont Object](Office.WebPageFont.md)



[WebPageFont Object Members](./overview/webpagefont-members-office.md)

