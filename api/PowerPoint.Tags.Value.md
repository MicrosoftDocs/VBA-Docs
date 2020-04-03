---
title: Tags.Value method (PowerPoint)
keywords: vbapp10.chm611009
f1_keywords:
- vbapp10.chm611009
ms.prod: powerpoint
api_name:
- PowerPoint.Tags.Value
ms.assetid: 8d7507d2-6533-5d63-c6ff-fec9581fb44f
ms.date: 06/08/2017
localization_priority: Normal
---


# Tags.Value method (PowerPoint)

Returns the value of the specified tag as a  **String**.


## Syntax

_expression_.**Value** (_Index_)

_expression_ A variable that represents a [Tags](PowerPoint.Tags.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The tag number.|

## Return value

String


## Example

This example displays the name and value for each tag associated with slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Tags

    For i = 1 To .Count

        MsgBox "Tag #" & i & ": Name = " & .Name(i)

        MsgBox "Tag #" & i & ": Value = " & .Value(i)

    Next

End With


```

This example searches through the tags for each slide in the active presentation. If there's a tag named "PRIORITY," a message box displays the tag value. If there isn't a tag named "PRIORITY," the example adds this tag that has the value "Unknown."




```vb
For Each s In Application.ActivePresentation.Slides
    With s.Tags
        found = False

        For i = 1 To .Count
            If .Name(i) = "PRIORITY" Then
                found = True
                slNum = .Parent.SlideIndex
                MsgBox "Slide " & slNum & " priority: " & .Value(i)
            End If
        Next

        If Not found Then
            slNum = .Parent.SlideIndex
            .Add "Name", "New Figures"
            .Add "Priority", "Unknown"
            MsgBox "Slide " & slNum & _
               " priority tag added: Unknown"
        End If
    End With
Next
```


## See also


[Tags Object](PowerPoint.Tags.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]