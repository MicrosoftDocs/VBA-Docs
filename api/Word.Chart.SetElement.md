---
title: Chart.SetElement method (Word)
ms.prod: word
api_name:
- Word.Chart.SetElement
ms.assetid: d172a9df-b081-0077-18ef-f75bf0d6f26a
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.SetElement method (Word)

Sets chart elements on a chart. Read/write  **MsoChartElementType**.


## Syntax

_expression_.**SetElement** (_Element_)

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Element_|Required| **MsoChartElementType**|One of the enumeration values that specifies the chart element type.|

## Remarks

For charts, the following commands in the  **Layout** tab correspond to the **SetElement** method:


- Everything in the  **Labels** group.
    
- Everything in the  **Axes** group.
    
- Everything in the  **Analysis** group.
    
-  **PlotArea**,  **Chart Wall**, and  **Chart Floor** buttons.
    


 **MsoChartElementType** is an enumeration of constants that refer to all of the above commands.


## Example

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


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]