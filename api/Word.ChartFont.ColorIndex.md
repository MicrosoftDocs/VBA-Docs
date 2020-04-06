---
title: ChartFont.ColorIndex property (Word)
keywords: vbawd10.chm255918086
f1_keywords:
- vbawd10.chm255918086
ms.prod: word
api_name:
- Word.ChartFont.ColorIndex
ms.assetid: 50f2415c-ff1f-cc16-371f-24f29373f96d
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartFont.ColorIndex property (Word)

Returns or sets the color of the font. Read/write  **Variant**.


## Syntax

_expression_.**ColorIndex**

_expression_ A variable that represents a '[ChartFont](Word.ChartFont.md)' object.


## Remarks

The color is specified as an index value into the current color palette, or as one of the following  **[XlColorIndex](Word.xlcolorindex.md)** constants:


-  **xlColorIndexAutomatic**
    
-  **xlColorIndexNone**
    

## Example

The following example changes the font color in the title of the first chart in the active document to red.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Title 
 ' Set the color to red. 
 .Font.ColorIndex = 3 
 End If 
 End With 
 End If 
End With
```


## See also


[ChartFont Object](Word.ChartFont.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]