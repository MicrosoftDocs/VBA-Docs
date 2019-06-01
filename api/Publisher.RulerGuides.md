---
title: RulerGuides object (Publisher)
keywords: vbapb10.chm786431
f1_keywords:
- vbapb10.chm786431
ms.prod: publisher
api_name:
- Publisher.RulerGuides
ms.assetid: c58d3cb2-8cf8-74fa-2bf4-a931dc95a26a
ms.date: 06/01/2019
localization_priority: Normal
---


# RulerGuides object (Publisher)

A collection of **[RulerGuide](Publisher.RulerGuide.md)** objects that represents a gridline used to align objects on a page.
 
## Remarks

Use the **Add** method to add ruler gridlines to the **RulerGuides** collection. 

Use the **Count** property to return the total number of ruler guides, horizontal and vertical, in the collection. 


## Example

This example creates horizontal ruler guides and vertical ruler guides every half inch on the first page of the active publication.

```vb
Sub SetRulerGuides() 
 Dim intCount As Integer 
 Dim intPos As Integer 
 With ActiveDocument.Pages(1).RulerGuides 
 For intCount = 1 To 16 
 intPos = intPos + 36 
 .Add Position:=intPos, Type:=pbRulerGuideTypeVertical 
 Next intCount 
 intPos = 0 
 For intCount = 1 To 21 
 intPos = intPos + 36 
 .Add Position:=intPos, Type:=pbRulerGuideTypeHorizontal 
 Next intCount 
 End With 
End Sub
```

<br/>

The following example uses the **Count** property to create a loop that deletes each of the ruler guides in the collection.

```vb
Sub RemoveAllGuides() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).RulerGuides 
 For intCount = 1 To .Count 
 .Item(1).Delete 
 Next intCount 
 End With 
End Sub
```


## Methods

- [Add](Publisher.RulerGuides.Add.md)

## Properties

- [Application](Publisher.RulerGuides.Application.md)
- [Count](Publisher.RulerGuides.Count.md)
- [Item](Publisher.RulerGuides.Item.md)
- [Parent](Publisher.RulerGuides.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]