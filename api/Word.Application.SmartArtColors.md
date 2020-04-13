---
title: Application.SmartArtColors property (Word)
keywords: vbawd10.chm158335459
f1_keywords:
- vbawd10.chm158335459
ms.prod: word
api_name:
- Word.Application.SmartArtColors
ms.assetid: e2cb12c4-3162-2327-9210-bd912dffa8e9
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SmartArtColors property (Word)

Returns a [SmartArtColors](Office.SmartArtColors.md) object that represents the set of color styles that are currently loaded in the application. Read-only.


## Syntax

_expression_. `SmartArtColors`

 _expression_ An expression that returns a '[Application](Word.Application.md)' object.


## Remarks

The set of colors represented by the **SmartArtColors** property correspond to the available color styles on the **Change Colors** button on the **Design tab** on the **SmartArt Tools** contextual tab in Word.


## Example

The following code example adds a SmartArt graphic to the active document and then sets the SmartArt graphic color to "Dark 2 Outline".


```vb
Dim myShape As Shape 
Dim mySmartArt As SmartArt 
 
Set myShape = ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts(1), 50, 50, 200, 200) 
Set mySmartArt = myShape.SmartArt 
 
mySmartArt.Color = Application.SmartArtColors(2) 

```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]