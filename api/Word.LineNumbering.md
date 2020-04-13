---
title: LineNumbering object (Word)
keywords: vbawd10.chm2418
f1_keywords:
- vbawd10.chm2418
ms.prod: word
api_name:
- Word.LineNumbering
ms.assetid: a2dd1278-c7dd-af4c-be32-1daded5556d6
ms.date: 06/08/2017
localization_priority: Normal
---


# LineNumbering object (Word)

Represents line numbers in the left margin or to the left of each newspaper-style column.


## Remarks

Use the **LineNumbering** property to return the **LineNumbering** object. The following example applies line numbering to the text in the first section of the active document.


```vb
With ActiveDocument.Sections(1).PageSetup.LineNumbering 
 .Active = True 
 .CountBy = 5 
 .RestartMode = wdRestartPage 
End With
```

The following example applies line numbering to the pages in the current section.




```vb
Selection.PageSetup.LineNumbering.Active = True
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]