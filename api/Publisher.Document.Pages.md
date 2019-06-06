---
title: Document.Pages property (Publisher)
keywords: vbapb10.chm196631
f1_keywords:
- vbapb10.chm196631
ms.prod: publisher
api_name:
- Publisher.Document.Pages
ms.assetid: 2bb3e529-a459-b37c-c9ae-4cc059954a63
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.Pages property (Publisher)

Returns a **[Pages](Publisher.Pages.md)** collection representing all the pages in the specified publication.


## Syntax

_expression_.**Pages**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Example

The following example returns the **Pages** collection of the active publication and reports how many pages there are.

```vb
Dim pgsTemp As Pages 
 
Set pgsTemp = ActiveDocument.Pages 
 
With pgsTemp 
 MsgBox "There are " & .Count _ 
 & " page(s) in the active publication." 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]