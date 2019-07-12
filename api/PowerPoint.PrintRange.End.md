---
title: PrintRange.End property (PowerPoint)
keywords: vbapp10.chm519004
f1_keywords:
- vbapp10.chm519004
ms.prod: powerpoint
api_name:
- PowerPoint.PrintRange.End
ms.assetid: 39f470c1-b469-3411-95e4-c6701487c498
ms.date: 06/08/2017
localization_priority: Normal
---


# PrintRange.End property (PowerPoint)

Returns the number of the last slide in the specified print range. Read-only.


## Syntax

_expression_.**End**

_expression_ A variable that represents an [PrintRange](PowerPoint.PrintRange.md) object.


## Return value

Long


## Example

This example displays a message that indicates the starting and ending slide numbers for print range one in the active presentation.


```vb
With ActivePresentation.PrintOptions.Ranges

    If .Count > 0 Then
        With .Item(1)
            MsgBox "Print range 1 starts on slide " & .Start & _
                " and ends on slide " & .End
        End With
    End If

End With
```


## See also


[PrintRange Object](PowerPoint.PrintRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]