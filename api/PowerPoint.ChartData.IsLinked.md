---
title: ChartData.IsLinked property (PowerPoint)
keywords: vbapp10.chm689003
f1_keywords:
- vbapp10.chm689003
ms.prod: powerpoint
api_name:
- PowerPoint.ChartData.IsLinked
ms.assetid: 038ed026-a14c-2c5c-3f2e-c931fa9840b0
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartData.IsLinked property (PowerPoint)

 **True** if the data for the chart is linked to an external Microsoft Excel workbook. Read-only **Boolean**.


## Syntax

_expression_. `IsLinked`

_expression_ A variable that represents a '[ChartData](PowerPoint.ChartData.md)' object.


## Remarks

Using the  **[BreakLink](PowerPoint.ChartData.BreakLink.md)** method to remove the link to an Excel workbook sets this property to **False**.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example verifies whether the data for the first chart in the active document is linked to an external Excel workbook. If the data for the chart is linked, the example then uses the  **BreakLink** method to remove the link. If the data for the chart is not linked, the example uses the **[Activate](PowerPoint.ChartData.Activate.md)** method to display the embedded data for the chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartData

            If .IsLinked Then

                .BreakLink

            Else

                .Activate

            End If

        End With

    End If

End With
```


## See also


[ChartData Object](PowerPoint.ChartData.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]