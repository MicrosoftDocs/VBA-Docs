---
title: ShapeRange.Callout property (Excel)
keywords: vbaxl10.chm640099
f1_keywords:
- vbaxl10.chm640099
ms.prod: excel
api_name:
- Excel.ShapeRange.Callout
ms.assetid: 15078411-7968-27ba-aa73-2c5d69220b08
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.Callout property (Excel)

Returns a **[CalloutFormat](Excel.CalloutFormat.md)** object that contains callout formatting properties for the specified shape. Applies to **ShapeRange** objects that represent line callouts. Read-only.


## Syntax

_expression_.**Callout**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Example

This example adds to _myDocument_ an oval and a callout that points to the oval. The callout text won't have a border, but it will have a vertical accent bar that separates the text from the callout line.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
 .AddShape msoShapeOval, 180, 200, 280, 130 
 With .AddCallout(msoCalloutTwo, 420, 170, 170, 40) 
 .TextFrame.Characters.Text = "My oval" 
 With .Callout 
 .Accent = True 
 .Border = False 
 End With 
 End With 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]