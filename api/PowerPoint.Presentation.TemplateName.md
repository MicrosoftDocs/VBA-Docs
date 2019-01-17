---
title: Presentation.TemplateName Property (PowerPoint)
keywords: vbapp10.chm583008
f1_keywords:
- vbapp10.chm583008
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.TemplateName
ms.assetid: 50cea27c-8181-eb32-20ae-88ae1f7ab34c
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.TemplateName Property (PowerPoint)

Returns the name of the design template associated with the specified presentation. Read-only.


## Syntax

 _expression_. `TemplateName`

 _expression_ A variable that represents a [Presentation](./PowerPoint.Presentation.md) object.


## Return value

String


## Remarks

The returned string includes the MS-DOS file name extension (for file types that are registered) but doesn't include the full path.


## Example

The following example applies the design template Professional.pot to the presentation Pres1.ppt if it is not already applied to it.


```vb
With Presentations("Pres1.ppt")
    If .TemplateName <> "Professional.pot" Then
        .ApplyTemplate "c:\program files\microsoft office" & _
            "\templates\presentation designs\Professional.pot"
    End If
End With
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]