---
title: Shape.Callout property (PowerPoint)
keywords: vbapp10.chm547018
f1_keywords:
- vbapp10.chm547018
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Callout
ms.assetid: 381f8eaa-f373-b1aa-6a53-4086d7e887d8
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Callout property (PowerPoint)

Returns a **[CalloutFormat](PowerPoint.CalloutFormat.md)** object that contains callout formatting properties for the specified shape. Applies to **Shape** or **ShapeRange** objects that represent line callouts. Read-only.


## Syntax

_expression_.**Callout**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

CalloutFormat


## Example

This example adds to _myDocument_ an oval and a callout that points to the oval. The callout text won't have a border, but it will have a vertical accent bar that separates the text from the callout line.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    .AddShape msoShapeOval, 180, 200, 280, 130

    With .AddCallout(msoCalloutTwo, 420, 170, 170, 40)

        .TextFrame.TextRange.Text = "My oval"

        With .Callout

            .Accent = True

            .Border = False

        End With

    End With

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]