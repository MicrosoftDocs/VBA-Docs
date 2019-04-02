---
title: Application.SmartArtLayouts property (Word)
keywords: vbawd10.chm158335457
f1_keywords:
- vbawd10.chm158335457
ms.prod: word
api_name:
- Word.Application.SmartArtLayouts
ms.assetid: dcbaf620-0865-8f2f-ef97-456edd0d70e3
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SmartArtLayouts property (Word)

Returns a [SmartArtLayouts](Office.SmartArtLayouts.md) object that represents the set of SmartArt layouts that are currently loaded in the application. Read-only.


## Syntax

_expression_. `SmartArtLayouts`

 _expression_ An expression that returns a '[Application](Word.Application.md)' object.


## Remarks

The set of layouts represented by the  **SmartArtLayouts** property correspond to the available layouts in the **Layouts** group on the **Design tab** on the **SmartArt Tools** contextual tab in Word.


## Example

The following code example adds a SmartArt graphic to the active document and then sets the SmartArt graphic layout to "Grouped List".


```vb
Dim myShape As Shape 
Dim mySmartArt As SmartArt 
 
Set myShape = ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts(1), 50, 50, 200, 200) 
Set mySmartArt = myShape.SmartArt 
 
mySmartArt.Layout = Application.SmartArtLayouts(15)
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]