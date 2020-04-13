---
title: TabStop object (Word)
keywords: vbawd10.chm2388
f1_keywords:
- vbawd10.chm2388
ms.prod: word
api_name:
- Word.TabStop
ms.assetid: 5290ae79-f728-24a8-6bb0-267072cd0288
ms.date: 06/08/2017
localization_priority: Normal
---


# TabStop object (Word)

Represents a single tab stop. The **TabStop** object is a member of the **[TabStops](Word.tabstops.md)** collection. The **TabStops** collection represents all the custom and default tab stops in a paragraph or group of paragraphs.


## Remarks

Use  **TabStops** (Index), where Index is the location of the tab stop (in points) or the index number, to return a single **TabStop** object. Tab stops are indexed numerically from left to right along the ruler. The following example removes the first custom tab stop from the selected paragraphs.


```vb
Selection.Paragraphs.TabStops(1).Clear
```

The following example adds a right-aligned tab stop positioned at 2 inches to the selected paragraphs.




```vb
Selection.Paragraphs.TabStops(InchesToPoints(2)) _ 
 .Alignment = wdAlignTabRight
```

Use the **Add** method to add a tab stop. The following example adds two tab stops to the selected paragraphs. The first tab stop is a left-aligned tab with a dotted tab leader positioned at 1 inch (72 points). The second tab stop is centered and is positioned at 2 inches.




```vb
With Selection.Paragraphs.TabStops 
 .Add Position:=InchesToPoints(1), _ 
 Leader:=wdTabLeaderDots, Alignment:=wdAlignTabLeft 
 .Add Position:=InchesToPoints(2), Alignment:=wdAlignTabCenter 
End With
```

You can also add a tab stop by specifying a location with the **TabStops** property. The following example adds a right-aligned tab stop positioned at 2 inches to the selected paragraphs.




```vb
Selection.Paragraphs.TabStops(InchesToPoints(2)) _ 
 .Alignment = wdAlignTabRight
```


> [!NOTE] 
>  Set the **DefaultTabStop** property to adjust the spacing of default tab stops.


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]