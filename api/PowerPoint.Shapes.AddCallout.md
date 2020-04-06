---
title: Shapes.AddCallout method (PowerPoint)
keywords: vbapp10.chm543005
f1_keywords:
- vbapp10.chm543005
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddCallout
ms.assetid: e4b468d7-793a-09ae-fcfc-6a73db93c90e
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddCallout method (PowerPoint)

Creates a borderless line callout. Returns a **[Shape](PowerPoint.Shape.md)** object that represents the new callout.


## Syntax

_expression_. `AddCallout`( `_Type_`, `_Left_`, `_Top_`, `_Width_`, `_Height_` )

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**[MsoCalloutType](Office.MsoCalloutType.md)**|The type of callout line.|
| _Left_|Required|**Single**|The position, measured in points, of the left edge of the callout's bounding box relative to the left edge of the slide.|
| _Top_|Required|**Single**|The position, measured in points, of the top edge of the callout's bounding box relative to the top edge of the slide.|
| _Width_|Required|**Single**| The width of the callout's bounding box, measured in points.|
| _Height_|Required|**Single**|The height of the callout's bounding box, measured in points.|

## Return value

Shape


## Remarks

You can insert a greater variety of callouts by using the  **[AddShape](PowerPoint.Shapes.AddShape.md)** method.


## Example

This example adds a borderless callout with a freely-rotating one-segment callout line to myDocument and then sets the callout angle to 30 degrees.


```vb
Sub NewCallout() 
    Dim sldOne As Slide 
    Set sldOne = ActivePresentation.Slides(1) 
    sldOne.Shapes.AddCallout(Type:=msoCalloutTwo, Left:=50, Top:=50, _ 
        Width:=200, Height:=100).Callout.Angle = msoCalloutAngle30 
End Sub
```


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]