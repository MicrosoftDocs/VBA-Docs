---
title: Font.NumberSpacing property (Word)
keywords: vbawd10.chm156369072
f1_keywords:
- vbawd10.chm156369072
ms.prod: word
api_name:
- Word.Font.NumberSpacing
ms.assetid: 468d47e9-9bda-dd6e-5a55-4a11b8ce351e
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.NumberSpacing property (Word)

Returns or sets the number spacing setting for a font. Read/write [WdNumberSpacing](Word.WdNumberSpacing.md).


## Syntax

 _expression_. `NumberSpacing`

 _expression_ An expression that returns a '[Font](Word.Font.md)' object.


## Remarks

OpenType fonts support a proportional and tabular figure feature to control number spacing. Proportional number spacing handles each number as having a different width. For example, "1" is displayed as narrower than "5". Tabular number spacing handles numbers as equal in width so that they align vertically, which increases the readability, especially for financial information.


## Example

The following code example sets the number spacing to proportional for the font in the active document.


```vb
ActiveDocument.Range.Font.NumberSpacing = wdNumberSpacingProportional
```


## See also


[Font Object](Word.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]