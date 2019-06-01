---
title: TabStop object (Publisher)
keywords: vbapb10.chm5701631
f1_keywords:
- vbapb10.chm5701631
ms.prod: publisher
api_name:
- Publisher.TabStop
ms.assetid: 74e71d75-503f-ef57-ddeb-24a788402df2
ms.date: 06/01/2019
localization_priority: Normal
---


# TabStop object (Publisher)

Represents a single tab stop. The **TabStop** object is a member of the **[TabStops](Publisher.TabStops.md)** collection. The **TabStops** collection represents all the custom and default tab stops in a paragraph or group of paragraphs.
 
## Remarks

Set the **[Document.DefaultTabStop](Publisher.Document.DefaultTabStop.md)** property to adjust the spacing of default tab stops.
 
Use **[Tabs](Publisher.ParagraphFormat.Tabs.md)** (_index_), where _index_ is the location of the tab stop (in [points](../language/glossary/vbe-glossary.md#point)) or the index number, to return a single **TabStop** object. Tab stops are indexed numerically from left to right along the ruler. 

Use the **[Add](Publisher.TabStops.Add.md)** method to add a tab stop. 

## Example

The following example removes the first custom tab stop from the selected paragraphs.

```vb
Sub ClearTabStop() 
 Selection.TextRange.ParagraphFormat.Tabs(1).Clear 
End Sub
```

<br/>

The following example adds a right-aligned tab stop positioned at 2 inches to the selected paragraphs.

```vb
Sub ChangeTabStop() 
 Selection.TextRange.ParagraphFormat.Tabs(2) _ 
 .Alignment = pbTabAlignmentTrailing 
End Sub
```

<br/>

The following example adds two tab stops to the selected paragraphs. The first tab stop is a left-aligned tab with a dotted tab leader positioned at 1 inch (72 points). The second tab stop is centered and is positioned at 2 inches.

```vb
Sub AddNewTabs() 
 With Selection.TextRange.ParagraphFormat.Tabs 
 .Add Position:=InchesToPoints(1), _ 
 Leader:=pbTabLeaderDot, Alignment:=pbTabAlignmentLeading 
 .Add Position:=InchesToPoints(2), _ 
 Leader:=pbTabLeaderNone, Alignment:=pbTabAlignmentCenter 
 End With 
End Sub
```


## Methods

- [Clear](Publisher.TabStop.Clear.md)

## Properties

- [Alignment](Publisher.TabStop.Alignment.md)
- [Application](Publisher.TabStop.Application.md)
- [Leader](Publisher.TabStop.Leader.md)
- [Parent](Publisher.TabStop.Parent.md)
- [Position](Publisher.TabStop.Position.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]