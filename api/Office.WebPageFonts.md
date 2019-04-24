---
title: WebPageFonts object (Office)
keywords: vbaof11.chm225000
f1_keywords:
- vbaof11.chm225000
ms.prod: office
api_name:
- Office.WebPageFonts
ms.assetid: c42bd65d-7c5c-148a-6f52-7aacd75be06a
ms.date: 01/29/2019
localization_priority: Normal
---


# WebPageFonts object (Office)

A collection of **[WebPageFont](office.webpagefont.md)** objects that describe the proportional font, proportional font size, fixed-width font, and fixed-width font size used when documents are saved as webpages. You can specify a different set of webpage font properties for each available character set.


## Remarks

The **WebPageFonts** collection contains one **WebPageFont** object for each character set.


## Example

The following example uses the **Item** property to set "myFont" to the **WebPageFont** object for the **English/Western European/Other Latin Script** character set in the current application.


```vb
Dim myFont As WebPageFont 
Set myFont = _ 
 Application.DefaultWebOptions.Fonts.Item_ 
 (msoCharacterSetEnglishWesternEuropeanOtherLatinScript)
```


## See also

- [WebPageFonts object members](overview/Library-Reference/webpagefonts-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]