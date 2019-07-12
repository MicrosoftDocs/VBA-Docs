---
title: LinkFormat.BreakLink method (PowerPoint)
keywords: vbapp10.chm563006
f1_keywords:
- vbapp10.chm563006
ms.prod: powerpoint
api_name:
- PowerPoint.LinkFormat.BreakLink
ms.assetid: cc177e67-8664-7273-2339-7d9c01f65ba6
ms.date: 06/08/2017
localization_priority: Normal
---


# LinkFormat.BreakLink method (PowerPoint)

Breaks the link between the source file and the specified OLE object, picture, or linked field.


## Syntax

_expression_. `BreakLink`

 _expression_ An expression that returns a [LinkFormat](PowerPoint.LinkFormat.md) object.


## Return value

Nothing


## Example

This example shows how to update and then break the links to any shapes that are linked to OLE objects on slide one in the active presentation.


```vb
Public Sub BreakLink_Example()



    Dim pptShape As Shape

    

    For Each pptShape In ActivePresentation.Slides(1).Shapes

        With pptShape

            If .Type = msoLinkedOLEObject Then

                .LinkFormat.Update

                .LinkFormat.BreakLink

            End If

        End With

    Next pptShape



End Sub
```


## See also


[LinkFormat Object](PowerPoint.LinkFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]