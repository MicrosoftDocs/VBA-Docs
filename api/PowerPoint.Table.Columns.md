---
title: Table.Columns property (PowerPoint)
keywords: vbapp10.chm622003
f1_keywords:
- vbapp10.chm622003
ms.prod: powerpoint
api_name:
- PowerPoint.Table.Columns
ms.assetid: 0645fa19-d5a2-1f4c-ae15-9623925d39bc
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.Columns property (PowerPoint)

Returns a **[Columns](PowerPoint.Columns.md)** collection that represents all the columns in a table. Read-only.


## Syntax

_expression_.**Columns**

_expression_ A variable that represents a [Table](PowerPoint.Table.md) object.


## Return value

Columns


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../powerpoint/How-to/return-objects-from-collections.md).


## Example

This example displays the shape number, the slide number, and the number of columns in the first table of the active presentation.


```vb
Dim ColCount As Integer

Dim sl As Integer

Dim sh As Integer



With ActivePresentation

    For sl = 1 To .Slides.Count
        For sh = 1 To .Slides(sl).Shapes.Count
            If .Slides(sl).Shapes(sh).HasTable Then
                ColCount = .Slides(sl).Shapes(sh) _
                    .Table.Columns.Count

                MsgBox "Shape " & sh & " on slide " & sl & _
                    " contains the first table and has " & _
                    ColCount & " columns."

                Exit Sub
            End If
        Next
    Next

End With
```

This example tests the selected shape to see if it contains a table. If it does, the code sets the width of column one to 72 points (one inch).




```vb
With ActiveWindow.Selection.ShapeRange

    If .HasTable = True Then

       .Table.Columns(1).Width = 72

    End If

End With
```


## See also


[Table Object](PowerPoint.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]