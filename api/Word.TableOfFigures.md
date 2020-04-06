---
title: TableOfFigures object (Word)
keywords: vbawd10.chm2337
f1_keywords:
- vbawd10.chm2337
ms.prod: word
api_name:
- Word.TableOfFigures
ms.assetid: 0a93d274-cd2e-3d65-48bc-b8fea84ff268
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfFigures object (Word)

Represents a single table of figures in a document. The  **TableOfFigures** object is a member of the **[TablesOfFigures](Word.tablesoffigures.md)** collection. The **TablesOfFigures** collection includes all the tables of figures in a document.


## Remarks

Use  **TablesOfFigures** (Index), where Index is the index number, to return a single **TableOfFigures** object. The index number represents the position of the table of figures in the document. The following example updates the page numbers of the items in the first table of figures in the active document.


```vb
ActiveDocument.TablesOfFigures(1).UpdatePageNumbers
```

Use the  **Add** method to add a table of figures to a document. A table of figures lists figure captions in the order in which they appear in the document. The following example replaces the selection in the active document with a table of figures that includes caption labels and page numbers.




```vb
ActiveDocument.TablesOfFigures.Add Range:=Selection.Range, _ 
 IncludeLabel:=True, IncludePageNumbers:=True
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]