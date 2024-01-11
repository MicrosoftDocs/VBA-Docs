---
title: Presentation.TemplateName property (PowerPoint)
keywords: vbapp10.chm583008
f1_keywords:
- vbapp10.chm583008
api_name:
- PowerPoint.Presentation.TemplateName
ms.assetid: 50cea27c-8181-eb32-20ae-88ae1f7ab34c
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Presentation.TemplateName property (PowerPoint)

Returns the name of the first design/master associated with the specified presentation. Read-only.


## Syntax

_expression_. `TemplateName`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

String


## Remarks

For pre-Office XML formats e.g. pot/ppt, the returned string previously returned the name of the template on which the presentation was created from, including the MS-DOS file name extension (for file types that are registered) but didn't include the full path. In the new XML formats e.g. potx/pptx, it now returns the name of the first design/master.


## Example

The following example applies the design template Professional.potx to the presentation Pres1.pptx if it is not already applied to it.


```vb
With Presentations("Pres1.pptx")
    If .TemplateName <> "Professional" Then
        .ApplyTemplate "c:\program files\microsoft office" & _
            "\templates\presentation designs\Professional.potx"
    End If
End With
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
