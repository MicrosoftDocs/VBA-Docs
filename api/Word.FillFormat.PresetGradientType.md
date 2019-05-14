---
title: FillFormat.PresetGradientType property (Word)
keywords: vbawd10.chm164102251
f1_keywords:
- vbawd10.chm164102251
ms.prod: word
api_name:
- Word.FillFormat.PresetGradientType
ms.assetid: b53ed5f8-61be-1abd-d3c7-e47a4ffc44b9
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.PresetGradientType property (Word)

Returns the preset gradient type for the specified fill. Read-only  **MsoPresetGradientType**.


## Syntax

_expression_.**PresetGradientType**

 _expression_ An expression that represents a **[FillFormat](word.fillformat.md)** object.


## Remarks

Use the  **[PresetGradient](Word.FillFormat.PresetGradient.md)** method to set the preset gradient type for the fill.


## Example

This example changes the fill for all shapes in _myDocument_ with the Moss preset gradient fill to the Fog preset gradient fill.


```vb
Set myDocument = ActiveDocument 
For Each s In myDocument.Shapes 
 With s.Fill 
 If .PresetGradientType = msoGradientMoss Then 
 .PresetGradient msoGradientHorizontal, 1, _ 
 msoGradientFog 
 End If 
 End With 
Next
```


## See also


[FillFormat Object](Word.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]