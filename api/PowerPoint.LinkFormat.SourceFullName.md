---
title: LinkFormat.SourceFullName property (PowerPoint)
keywords: vbapp10.chm563003
f1_keywords:
- vbapp10.chm563003
ms.prod: powerpoint
api_name:
- PowerPoint.LinkFormat.SourceFullName
ms.assetid: 6a7fb694-609a-77c5-eabc-d95693a87299
ms.date: 06/08/2017
localization_priority: Normal
---


# LinkFormat.SourceFullName property (PowerPoint)

Returns or sets the name and path of the source file for the linked OLE object. Read/write.


## Syntax

_expression_. `SourceFullName`

_expression_ A variable that represents a [LinkFormat](PowerPoint.LinkFormat.md) object.


## Return value

String


## Example

This example sets the source file for shape one on slide one in the active presentation to Wordtest.doc and specifies that the object's image be updated automatically.


```vb
With ActivePresentation.Slides(1).Shapes(1)

    If .Type = msoLinkedOLEObject Then

        With .LinkFormat

            .SourceFullName = "c:\my documents\wordtest.doc"

            .AutoUpdate = ppUpdateOptionAutomatic

        End With

    End If

End With
```


## See also


[LinkFormat Object](PowerPoint.LinkFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]