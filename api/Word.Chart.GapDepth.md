---
title: Chart.GapDepth property (Word)
ms.prod: word
api_name:
- Word.Chart.GapDepth
ms.assetid: 09147a74-c8bb-4fc5-0389-c8f46e0be67d
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.GapDepth property (Word)

Returns or sets the distance, as a percentage of the marker width, between the data series in a 3D chart. Read/write  **Long**.


## Syntax

_expression_.**GapDepth**

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Remarks

The value of this property must be between 0 and 500. 


> [!NOTE] 
> This property applies only to 3D charts.


## Example

The following example sets the distance between the data series for the first chart in the active document to 200 percent of the marker width. You should run the example on a 3D chart (the  **GapDepth** property fails on 2D charts).


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.GapDepth = 200 
 End If 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]