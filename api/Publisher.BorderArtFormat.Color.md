---
title: BorderArtFormat.Color property (Publisher)
keywords: vbapb10.chm7602183
f1_keywords:
- vbapb10.chm7602183
ms.prod: publisher
api_name:
- Publisher.BorderArtFormat.Color
ms.assetid: fb2fe2f7-d321-43d3-232d-db3b513dae43
ms.date: 06/05/2019
localization_priority: Normal
---


# BorderArtFormat.Color property (Publisher)

Returns a **[ColorFormat](Publisher.ColorFormat.md)** object representing the color information for the specified object.


## Syntax

_expression_.**Color**

_expression_ A variable that represents a **[BorderArtFormat](Publisher.BorderArtFormat.md)** object.


## Example

This example tests the font color of the first story in the active document and tells the user if the font color is black or not.

```vb
Sub FontColor() 
 
 If Application.ActiveDocument.Stories(1) _ 
 .TextRange.Font.Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) Then 
 MsgBox "Your font color is black" 
 Else 
 MsgBox "Your font color is not black" 
 End If 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]