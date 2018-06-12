---
title: Border Object (Word)
keywords: vbawd10.chm2363
f1_keywords:
- vbawd10.chm2363
ms.prod: word
api_name:
- Word.Border
ms.assetid: be48c020-b86c-c004-ce1c-76d9edae9791
ms.date: 06/08/2017
---


# Border Object (Word)

Represents a border of an object. The  **Border** object is a member of the **[Borders](Word.borders.md)** collection.


## Remarks

Use  **Borders** (index), where index identifies the border, to return a single **Border** object. Index can be one of the **[WdBorderType](Word.WdBorderType.md)** constants. Some of the **WdBorderType** constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.

Use the  **[LineStyle](Word.Border.LineStyle.md)** property to apply a border line to a **Border** object. The following example applies a double-line border below the first paragraph in the active document.




```
With ActiveDocument.Paragraphs(1).Borders(wdBorderBottom) 
 .LineStyle = wdLineStyleDouble 
 .LineWidth = wdLineWidth025pt 
End With
```

The following example applies a single-line border around the first character in the selection.




```
With Selection.Characters(1) 
 .Font.Size = 36 
 .Borders.Enable = True 
End With
```

The following example adds an art border around each page in the first section.




```
For Each aBorder In ActiveDocument.Sections(1).Borders 
 With aBorder 
 .ArtStyle = wdArtSeattle 
 .ArtWidth = 20 
 End With 
Next aBorder
```

 **Border** objects cannot be added to the **[Borders](Word.borders.md)** collection. The number of members in the **[Borders](Word.borders.md)** collection is finite and varies depending on the type of object. For example, a table has six elements in the **[Borders](Word.borders.md)** collection, whereas a paragraph has four.


## Properties



|**Name**|
|:-----|
|[Application](Word.Border.Application.md)|
|[ArtStyle](Word.Border.ArtStyle.md)|
|[ArtWidth](Word.Border.ArtWidth.md)|
|[Color](Word.Border.Color.md)|
|[ColorIndex](Word.Border.ColorIndex.md)|
|[Creator](Word.Border.Creator.md)|
|[Inside](Word.Border.Inside.md)|
|[LineStyle](Word.Border.LineStyle.md)|
|[LineWidth](border-linewidth-property-word.md)|
|[Parent](border-parent-property-word.md)|
|[Visible](border-visible-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
