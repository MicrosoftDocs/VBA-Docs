---
title: TextEffectFormat.FontBold property (Word)
keywords: vbawd10.chm164560997
f1_keywords:
- vbawd10.chm164560997
ms.prod: word
api_name:
- Word.TextEffectFormat.FontBold
ms.assetid: 7432680f-5dbd-ae1c-3d49-ee99cd9f93bb
ms.date: 06/08/2017
localization_priority: Normal
---


# TextEffectFormat.FontBold property (Word)

Sets the font to bold for the specified Word Art shape. Read/write  **MsoTriState**.


## Syntax

_expression_.**FontBold**

_expression_ A variable that represents a '[TextEffectFormat](Word.TextEffectFormat.md)' object.


## Example

This example sets the font to bold for the third shape on the active document if the shape is WordArt.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
With docActive.Shapes(3) 
 If .Type = msoTextEffect Then 
 .TextEffect.FontBold = msoTrue 
 End If 
End With
```


## See also


[TextEffectFormat Object](Word.TextEffectFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]