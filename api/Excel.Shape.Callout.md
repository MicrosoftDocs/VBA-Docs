---
title: Shape.Callout property (Excel)
keywords: vbaxl10.chm636092
f1_keywords:
- vbaxl10.chm636092
ms.prod: excel
api_name:
- Excel.Shape.Callout
ms.assetid: 80c67ea9-7e55-9841-bbed-302cbd669ce5
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.Callout property (Excel)

Returns a **[CalloutFormat](Excel.CalloutFormat.md)** object that contains callout formatting properties for the specified shape. Applies to a **Shape** object that represent line callouts. Read-only.


## Syntax

_expression_.**Callout**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


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