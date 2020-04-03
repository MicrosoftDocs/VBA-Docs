---
title: PrintRange object (PowerPoint)
keywords: vbapp10.chm519000
f1_keywords:
- vbapp10.chm519000
ms.prod: powerpoint
api_name:
- PowerPoint.PrintRange
ms.assetid: 62f098b3-5e67-8fa4-3af9-4507160fa1ad
ms.date: 06/08/2017
localization_priority: Normal
---


# PrintRange object (PowerPoint)

Represents a single range of consecutive slides or pages to be printed.


## Remarks

 The **PrintRange** object is a member of the **[PrintRanges](PowerPoint.PrintRanges.md)** collection. The **PrintRanges** collection contains all the print ranges that have been defined for the specified presentation.

You can set print ranges in the  **PrintRanges** collection independent of the **RangeType** setting; these ranges are retained as long as the presentation they're contained in is loaded. The ranges in the **PrintRanges** collection are applied when the **RangeType** property is set to **ppPrintSlideRange**.


## Example

Use  **Ranges** (_index_), where _index_ is the print range index number, to return a single **PrintRange** object. The following example displays a message that indicates the starting and ending slide numbers for print range one in the active presentation.


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

Use the [Add](PowerPoint.PrintRanges.Add.md)method to create a **PrintRange** object and add it to the **PrintRanges** collection. The following example defines three print ranges that represent slide 1, slides 3 through 5, and slides 8 and 9 in the active presentation and then prints the slides in these ranges.




```vb
With ActivePresentation.PrintOptions

    .RangeType = ppPrintSlideRange

    With .Ranges

        .ClearAll

        .Add 1, 1

        .Add 3, 5

        .Add 8, 9

    End With

End With

ActivePresentation.PrintOut
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]