---
title: FindReplace.Execute method (Publisher)
keywords: vbapb10.chm8323086
f1_keywords:
- vbapb10.chm8323086
ms.prod: publisher
api_name:
- Publisher.FindReplace.Execute
ms.assetid: 351a64ab-3c6c-c9c9-7ffe-b60b73d390ae
ms.date: 06/07/2019
localization_priority: Normal
---


# FindReplace.Execute method (Publisher)

Performs the specified find or replace operation.


## Syntax

_expression_.**Execute**

_expression_ A variable that represents a **[FindReplace](Publisher.FindReplace.md)** object.


## Return value

Boolean


## Example

This example executes a find and replace operation on the active document.

```vb
Sub ExecuteFindReplace() 
 Dim objFindReplace As FindReplace 
 Set objFindReplace = ActiveDocument.Find 
 With objFindReplace 
 .Clear 
 .FindText = "library" 
 .Execute 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]