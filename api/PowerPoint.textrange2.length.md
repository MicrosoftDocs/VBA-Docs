---
title: TextRange2.Length property (PowerPoint)
ms.assetid: a9e3fa70-7bca-460d-8d5d-26f844b47d9b
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# TextRange2.Length property (PowerPoint)

Get a Long that represents the length of a text range. Read-only.


## Syntax

_expression_.**Length**

 _expression_ An expression that returns a 'TextRange2' object.


## Return value

Long


## Example

This example adds a shape with text and rotates the shape without rotating the text in the active PowerPoint presentation.


```vb
Sub SetTextRange() 
 Dim shpStar As Shape 
 Dim sldOne As Slide 
 Dim effNew As Effect 
 
 Set sldOne = ActivePresentation.Slides(1) 
 Set shpStar = sldOne.Shapes.AddShape(Type:=msoShape5pointStar, _ 
 Left:=32, Top:=32, Width:=300, Height:=300) 
 
 shpStar.TextFrame.TextRange2.Text = "Animated shape." 
 
 Set effNew = sldOne.TimeLine.MainSequence.AddEffect(Shape:=shpStar, _ 
 EffectId:=msoAnimEffectPath5PointStar, Level:=msoAnimateTextByAllLevels, _ 
 Trigger:=msoAnimTriggerAfterPrevious) 
 With effNew 
 If .TextRangeStart = 0 And .TextRangeLength > 0 Then 
 With .Behaviors.Add(Type:=msoAnimTypeRotation).RotationEffect 
 .From = 0 
 .To = 360 
 End With 
 .Timing.AutoReverse = msoTrue 
 End If 
 End With 
End Sub
```


## See also


[TextRange2 object (PowerPoint)](PowerPoint.textrange2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]