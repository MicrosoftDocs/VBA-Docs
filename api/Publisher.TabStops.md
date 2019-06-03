---
title: TabStops object (Publisher)
keywords: vbapb10.chm5636095
f1_keywords:
- vbapb10.chm5636095
ms.prod: publisher
api_name:
- Publisher.TabStops
ms.assetid: fbaa194c-754a-3437-c3d5-fa70c951ca4f
ms.date: 06/01/2019
localization_priority: Normal
---


# TabStops object (Publisher)

A collection of **[TabStop](Publisher.TabStop.md)** objects that represent the custom and default tabs for a paragraph or group of paragraphs.
 
## Remarks

Use the **[ParagraphFormat.Tabs](Publisher.ParagraphFormat.Tabs.md)** property to return the **TabStops** collection. Use **Tabs** (_index_), where _index_ is the location of the tab stop (in [points](../language/glossary/vbe-glossary.md#point)) or the index number, to return a single **TabStop** object. Tab stops are indexed numerically from left to right along the ruler. 

Use the **Add** method to add a tab stop. 



## Example

The following example clears all the custom tab stops from the first paragraph in the active publication.

```vb
Sub ClearAllTabStops() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .ParagraphFormat.Tabs.ClearAll 
End Sub
```

<br/>

The following example adds a tab stop positioned at 2.5 inches to the selected paragraphs and then displays the position of each item in the **TabStops** collection.

```vb
Sub Tabs() 
 Dim intTab As Integer 
 Selection.TextRange.ParagraphFormat.Tabs _ 
 .Add Position:=InchesToPoints(2.5), _ 
 Alignment:=pbTabAlignmentLeading, Leader:=pbTabLeaderNone 
 With Selection.TextRange.ParagraphFormat 
 For intTab = 1 To .Tabs.Count 
 MsgBox "Position = " & PointsToInches _ 
 (.Tabs(intTab).Position) & " inches" 
 intTab = intTab + 1 
 Next intTab 
 End With 
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

<br/>

The following example removes the first custom tab stop from the first paragraph in the active publication.

```vb
Sub ClearTabStop() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .ParagraphFormat.Tabs(1).Clear 
End Sub
```

<br/>

The following example changes the second tab in the selection to a right-aligned tab stop.

```vb
Sub ChangeTabStop() 
 Selection.TextRange.ParagraphFormat.Tabs(2) _ 
 .Alignment = pbTabAlignmentTrailing 
End Sub
```


## Methods

- [Add](Publisher.TabStops.Add.md)
- [ClearAll](Publisher.TabStops.ClearAll.md)
- [Item](Publisher.TabStops.Item.md)

## Properties

- [Application](Publisher.TabStops.Application.md)
- [Count](Publisher.TabStops.Count.md)
- [Parent](Publisher.TabStops.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]