---
title: Presentation.Tags property (PowerPoint)
keywords: vbapp10.chm583018
f1_keywords:
- vbapp10.chm583018
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Tags
ms.assetid: 3b75d7ae-ce76-0023-c11e-1f39f4319ed5
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.Tags property (PowerPoint)

Returns a **[Tags](PowerPoint.Tags.md)** object that represents the tags for the specified object. Read-only.


## Syntax

_expression_. `Tags`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

Tags


## Example


> [!NOTE] 
> Tag values are added and stored in uppercase text. You should perform tests on tag values using uppercase text, as shown in the second example.

This example adds a tag named "REGION" and a tag named "PRIORITY" to slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Tags

    .Add "Region", "East"     'Adds "Region" tag with value "East"

    .Add "Priority", "Low"    'Adds "Priority" tag with value "Low"

End With
```

This example searches through the tags for each slide in the active presentation. If there's a tag named "PRIORITY," a message box displays the tag value. If the object doesn't have a tag named "PRIORITY," the example adds this tag with the value "Unknown."




```vb
For Each s In Application.ActivePresentation.Slides

    With s.Tags

        found = False

        For i = 1 To .Count

          If .Name(i) = "PRIORITY" Then

              found = True

              slNum = .Parent.SlideIndex

              MsgBox "Slide " & slNum & " Priority: " & .Value(i)

          End If

        Next

        If Not found Then

          slNum = .Parent.SlideIndex

          .Add "Priority", "Unknown"

          MsgBox "Slide " & slNum & " Priority tag added: Unknown"

        End If

    End With

Next
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]