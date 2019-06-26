---
title: TableOfAuthorities object (Word)
keywords: vbawd10.chm2321
f1_keywords:
- vbawd10.chm2321
ms.prod: word
api_name:
- Word.TableOfAuthorities
ms.assetid: abd7d600-8b20-0752-4629-8a4f5193dd5d
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfAuthorities object (Word)

Represents a single table of authorities in a document (a TOA field). The  **TableOfAuthorities** object is a member of the **[TablesOfAuthorities](Word.tablesofauthorities.md)** collection. The **TablesOfAuthorities** collection includes all the tables of authorities in a document.


## Remarks

Use  **TablesOfAuthorities** (Index), where Index is the index number, to return a single **TableOfAuthorities** object. The index number represents the position of the table of authorities in the document. The following example includes category headers in the first table of authorities in the active document and then updates the table.


```vb
With ActiveDocument.TablesOfAuthorities(1) 
 .IncludeCategoryHeader = True 
 .Update 
End With
```

Use the  **Add** method to add a table of authorities to a document. The following example adds a table of authorities that includes all categories at the beginning of the active document.




```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.TablesOfAuthorities.Add Range:=myRange, _ 
 Passim:=True, Category:=0, EntrySeparator:=", "
```


> [!NOTE] 
> A table of authorities is built from TA (Table of Authorities Entry) fields in a document. Use the  **MarkCitation** method to mark citations to be included in a table of authorities.


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]