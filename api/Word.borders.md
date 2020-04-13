---
title: Borders object (Word)
ms.prod: word
ms.assetid: 6dd1d4cc-2dcf-22c7-a299-4721a5543ba3
ms.date: 06/08/2017
localization_priority: Normal
---


# Borders object (Word)

A collection of  **[Border](Word.Border.md)** objects that represent the borders of an object.


## Remarks

Use the **Borders** property to return the **Borders** collection. The following example applies the default border around the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).Borders.Enable = True
```

**Border** objects cannot be added to the **Borders** collection. The number of members in the **Borders** collection is finite and varies depending on the type of object. For example, a table has six elements in the **Borders** collection, whereas a paragraph has four.

Use  **Borders** (index), where index identifies the border, to return a single **Border** object. Index can be one of the **[WdBorderType](Word.WdBorderType.md)** constants. Some of the **WdBorderType** constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.

Use the **[LineStyle](Word.Border.LineStyle.md)** property to apply a border line to a **Border** object. The following example applies a double-line border below the first paragraph in the active document.




```vb
With ActiveDocument.Paragraphs(1).Borders(wdBorderBottom) 
 .LineStyle = wdLineStyleDouble 
 .LineWidth = wdLineWidth025pt 
End With
```

The following example applies a single-line border around the first character in the selection.




```vb
With Selection.Characters(1) 
 .Font.Size = 36 
 .Borders.Enable = True 
End With
```

The following example adds an art border around each page in the first section.




```vb
For Each aBorder In ActiveDocument.Sections(1).Borders 
 With aBorder 
 .ArtStyle = wdArtSeattle 
 .ArtWidth = 20 
 End With 
Next aBorder
```


## Methods

- [ApplyPageBordersToAllSections](Word.Borders.ApplyPageBordersToAllSections.md)
- [Item](Word.Borders.Item.md)

## Properties

- [AlwaysInFront](Word.Borders.AlwaysInFront.md)
- [Application](Word.Borders.Application.md)
- [Count](Word.Borders.Count.md)
- [Creator](Word.Borders.Creator.md)
- [DistanceFrom](Word.Borders.DistanceFrom.md)
- [DistanceFromBottom](Word.Borders.DistanceFromBottom.md)
- [DistanceFromLeft](Word.Borders.DistanceFromLeft.md)
- [DistanceFromRight](Word.Borders.DistanceFromRight.md)
- [DistanceFromTop](Word.Borders.DistanceFromTop.md)
- [Enable](Word.Borders.Enable.md)
- [EnableFirstPageInSection](Word.Borders.EnableFirstPageInSection.md)
- [EnableOtherPagesInSection](Word.Borders.EnableOtherPagesInSection.md)
- [HasHorizontal](Word.Borders.HasHorizontal.md)
- [HasVertical](Word.Borders.HasVertical.md)
- [InsideColor](Word.Borders.InsideColor.md)
- [InsideColorIndex](Word.Borders.InsideColorIndex.md)
- [InsideLineStyle](Word.Borders.InsideLineStyle.md)
- [InsideLineWidth](Word.Borders.InsideLineWidth.md)
- [JoinBorders](Word.Borders.JoinBorders.md)
- [OutsideColor](Word.Borders.OutsideColor.md)
- [OutsideColorIndex](Word.Borders.OutsideColorIndex.md)
- [OutsideLineStyle](Word.Borders.OutsideLineStyle.md)
- [OutsideLineWidth](Word.Borders.OutsideLineWidth.md)
- [Parent](Word.Borders.Parent.md)
- [Shadow](Word.Borders.Shadow.md)
- [SurroundFooter](Word.Borders.SurroundFooter.md)
- [SurroundHeader](Word.Borders.SurroundHeader.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]