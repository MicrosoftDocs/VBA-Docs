---
title: DefaultWebOptions.Fonts property (Word)
keywords: vbawd10.chm165871631
f1_keywords:
- vbawd10.chm165871631
ms.prod: word
api_name:
- Word.DefaultWebOptions.Fonts
ms.assetid: 3a3247af-ae74-33f1-2944-0371bf13be6f
ms.date: 06/08/2017
localization_priority: Normal
---


# DefaultWebOptions.Fonts property (Word)

Returns the **WebPageFonts** collection representing the set of fonts that Microsoft Word uses when you open a webpage in Word.


## Syntax

_expression_.**Fonts**

_expression_ An expression that returns a **[DefaultWebOptions](Word.DefaultWebOptions.md)** object.


## Remarks

Word uses the fonts in the **WebPageFonts** collection to display webpages in Word when either there is no font information specified on the webpage or Word is unable to display the character set.


## Example

The following example sets the default fixed-width font for the English/Western European/Other Latin Script character set to Courier New, 14 point.

```vb
With Application.DefaultWebOptions _ 
 .Fonts(msoCharacterSetEnglishWesternEuropeanOtherLatinScript) 
 .FixedWidthFont = "Courier New" 
 .FixedWidthFontSize = 14 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]