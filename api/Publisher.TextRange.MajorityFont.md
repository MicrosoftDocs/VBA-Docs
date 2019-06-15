---
title: TextRange.MajorityFont property (Publisher)
keywords: vbapb10.chm5308467
f1_keywords:
- vbapb10.chm5308467
ms.prod: publisher
api_name:
- Publisher.TextRange.MajorityFont
ms.assetid: b0007ebc-ed0b-aab8-49fe-76353efbc1d2
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.MajorityFont property (Publisher)

Returns a **[Font](Publisher.Font.md)** object that represents the font name most in use in a text range.


## Syntax

_expression_.**MajorityFont**

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Return value

Font


## Example

This example creates a new text box, fills it with text, checks if the font most in use is Tahoma, and if it isn't, changes the font to Tahoma.

```vb
Sub SetFontName() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=100, Height:=100) _ 
 .TextFrame.TextRange 
 For intCount = 1 To 10 
 .InsertAfter NewText:="This is a test. " 
 Next intCount 
 If .MajorityFont <> "Tahoma" Then _ 
 .Font.Name = "Tahoma" 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]