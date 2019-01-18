---
title: Series.MarkerStyle Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Series.MarkerStyle
ms.assetid: e985978e-f0cf-b809-ebe1-f5504e9e8df6
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.MarkerStyle Property (PowerPoint)

Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write  **[xlMarkerStyle](PowerPoint.XlMarkerStyle.md)**.


## Syntax

 _expression_. `MarkerStyle`

 _expression_ A variable that represents a '[Series](PowerPoint.Series.md)' object.


## Remarks

 **MarkerStyle** can be one of the following **xlMarkerStyle** constants:


-  **xlMarkerStyleAutomatic** —Automatic markers.
    
-  **xlMarkerStyleCircle** —Circular markers.
    
-  **xlMarkerStyleDash** —Long bar markers.
    
-  **xlMarkerStyleDiamond** —Diamond-shaped markers.
    
-  **xlMarkerStyleDot** —Short bar markers.
    
-  **xlMarkerStyleNone** —No markers.
    
-  **xlMarkerStylePicture** —Picture markers.
    
-  **xlMarkerStylePlus** —Square markers with a plus sign.
    
-  **xlMarkerStyleSquare** —Square markers.
    
-  **xlMarkerStyleStar** —Square markers with an asterisk.
    
-  **xlMarkerStyleTriangle** —Triangular markers.
    
-  **xlMarkerStyleX** —Square markers with an X.
    



## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the marker style for series one for the first chart in the active document. You should run the example on a 2-D line chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle

    End If

End With
```


## See also


[Series Object](PowerPoint.Series.md)

