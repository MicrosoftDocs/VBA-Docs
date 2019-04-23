---
title: Application.SmartArtQuickStyles property (Word)
keywords: vbawd10.chm158335458
f1_keywords:
- vbawd10.chm158335458
ms.prod: word
api_name:
- Word.Application.SmartArtQuickStyles
ms.assetid: 47cca923-fc88-6973-926c-2fa69c2f0f10
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SmartArtQuickStyles property (Word)

Returns a [SmartArtQuickStyles](Office.SmartArtQuickStyles.md) object that represents the set of SmartArt styles that are currently loaded in the application. Read-only.


## Syntax

_expression_. `SmartArtQuickStyles`

 _expression_ An expression that returns a '[Application](Word.Application.md)' object.


## Remarks

The set of styles represented by the  **SmartArtQuickStyles** property correspond to the available styles in the **Styles** group on the **Design tab** on the **SmartArt Tools** contextual tab in Word.


## Example

The following code example adds a SmartArt graphic to the active document and then sets the SmartArt graphic style to "Polished".


```vb
Dim myShape As Shape 
Dim mySmartArt As SmartArt 
 
Set myShape = ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts(1), 50, 50, 200, 200) 
Set mySmartArt = myShape.SmartArt 
 
mySmartArt.QuickStyle = Application.SmartArtQuickStyles.Item(6)
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]