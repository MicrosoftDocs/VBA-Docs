---
title: StyleSheets object (Word)
keywords: vbawd10.chm3198
f1_keywords:
- vbawd10.chm3198
ms.prod: word
api_name:
- Word.StyleSheets
ms.assetid: 4b3da354-38a6-5758-3080-82a940b668c9
ms.date: 06/08/2017
localization_priority: Normal
---


# StyleSheets object (Word)

A collection of  **StyleSheet** objects that represents the cascading style sheets attached to a document. The **StyleSheets** collection includes all cascading style sheets displayed in the **Linked CSS Style Sheets** dialog box.


## Remarks

Use the **StyleSheets** property to return the **StyleSheets** collection. Use the **Add** method to add a style sheet to the **StyleSheets** collection. The following example adds three cascading style sheets to the active document and sets the third as the highest in precedence.


```vb
Sub AddCSS() 
 With ActiveDocument.StyleSheets 
 .Add FileName:="Web.css", Title:="Web Styles" 
 .Add FileName:="New.css", Linktype:=wdStyleSheetLinkTypeImported, _ 
 Title:="New Styles" 
 .Add FileName:="Defs.css", Title:="Definitions", _ 
 Precedence:=wdStyleSheetPrecedenceHighest 
 End With 
End Sub
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]