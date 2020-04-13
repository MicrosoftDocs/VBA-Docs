---
title: DataLabel object (Word)
keywords: vbawd10.chm3569
f1_keywords:
- vbawd10.chm3569
ms.prod: word
api_name:
- Word.DataLabel
ms.assetid: b955596d-ac94-1e18-4e72-cdf090fc1f9e
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabel object (Word)

Represents the data label on a chart point or trendline.


## Remarks

 On a series, the **DataLabel** object is a member of the **[DataLabels](Word.DataLabels.md)** collection. The **DataLabels** collection contains a **DataLabel** object for each point. For a series without definable points (such as an area series), the **DataLabels** collection contains a single **DataLabel** object.


## Example

Use  **[DataLabels](Word.Series.DataLabels.md)** (_index_), where _index_ is the data label index number, to return a single **DataLabel** object. The following example sets the number format for the fifth data label in the first series of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels(5).NumberFormat = "0.000" 
 End If 
End With 

```

Use the **[Point.DataLabel](Word.Point.DataLabel.md)** property to return the **DataLabel** object for a single point. The following example turns on the data label for the second point in the first series of the first chart in the active document and sets the data label text to "Saturday."




```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(1).Points(2) 
 .HasDataLabel = True 
 .DataLabel.Text = "Saturday" 
 End With 
 End If 
End With 

```

On a trendline, the **[Trendline.DataLabel](Word.Trendline.DataLabel.md)** property returns the text shown with the trendline. This can be the equation, the R-squared value, or both (if both are showing). The following example sets the trendline text for the first trendline in the first series of the first chart in the active document to show only the equation.




```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(1).Trendlines(1) 
 .DisplayRSquared = False 
 .DisplayEquation = True 
 End With 
 End If 
End With
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]