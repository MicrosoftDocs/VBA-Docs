---
title: Chart.SetElement method (PowerPoint)
keywords: vbapp10.chm684044
f1_keywords:
- vbapp10.chm684044
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.SetElement
ms.assetid: c9f05df8-a85a-c401-c9bc-33f3a2cc4561
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.SetElement method (PowerPoint)

Sets chart elements on a chart. Read/write  **MsoChartElementType**.


## Syntax

_expression_.**SetElement** (_Element_)

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Element_|Required|**MsoChartElementType**|One of the enumeration values that specifies the chart element type.|

## Remarks

For charts, the following commands in the  **Layout** tab correspond to the **SetElement** method:


- Everything in the  **Labels** group.
    
- Everything in the  **Axes** group.
    
- Everything in the  **Analysis** group.
    
- **PlotArea**,  **Chart Wall**, and  **Chart Floor** buttons.
    


 **MsoChartElementType** is an enumeration of constants that refer to all of the above commands.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets chart elements by using the various constant values to an active chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart

            ' Select the major gridlines on the value axis.

            .Axes(xlValue).MajorGridlines.Select

            .SetElement msoElementChartTitleCenteredOverlay

            .SetElement msoElementPrimaryCategoryGridLinesMinor

            ' Select the walls.

            .Walls.Select

            .SetElement msoElementChartFloorShow

        End With

    End If

End With


```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]