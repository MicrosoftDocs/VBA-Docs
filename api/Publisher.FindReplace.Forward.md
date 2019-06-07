---
title: FindReplace.Forward property (Publisher)
keywords: vbapb10.chm8323078
f1_keywords:
- vbapb10.chm8323078
ms.prod: publisher
api_name:
- Publisher.FindReplace.Forward
ms.assetid: a1a0046c-81be-62d6-8739-5dc843d249bc
ms.date: 06/07/2019
localization_priority: Normal
---


# FindReplace.Forward property (Publisher)

Sets or retrieves a **Boolean** representing the direction of the text search. **True** if the find operation searches forward through the document. **False** if it searches backward through the document. Read/write.


## Syntax

_expression_.**Forward**

_expression_ A variable that represents a **[FindReplace](Publisher.FindReplace.md)** object.


## Return value

Boolean


## Remarks

Forward must be set to **True** when replacing text.


## Example

This example replaces all occurrences of the word This in the selection with the word That in each open publication.

```vb
Dim objDocument As Document 
For Each objDocument In Documents 
 With objDocument.Find 
 .Clear 
 .MatchCase = True 
 .FindText = "This" 
 .ReplaceWithText = "That" 
 .ReplaceScope = pbReplaceScopeAll 
 .Forward = True 
 .Execute 
 End With 
Next objDocument 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]