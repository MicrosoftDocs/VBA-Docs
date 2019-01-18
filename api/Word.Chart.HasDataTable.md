---
title: Chart.HasDataTable property (Word)
keywords: vbawd10.chm79365492
f1_keywords:
- vbawd10.chm79365492
ms.prod: word
api_name:
- Word.Chart.HasDataTable
ms.assetid: 62af9540-9a69-0e19-b884-4f2b5947152f
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.HasDataTable property (Word)

 **True** if the chart has a data table. Read/write **Boolean**.


## Syntax

 _expression_. `HasDataTable`

 _expression_ A variable that represents a '[Chart](Word.Chart.md)' object.


## Example

The following example causes the embedded chart data table to be displayed with an outline border and no cell borders.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .HasDataTable = True 
 With .DataTable 
 .HasBorderHorizontal = False 
 .HasBorderVertical = False 
 .HasBorderOutline = True 
 End With 
 End With 
 End If 
End With 

```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]