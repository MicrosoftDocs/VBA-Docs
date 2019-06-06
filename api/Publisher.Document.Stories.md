---
title: Document.Stories property (Publisher)
keywords: vbapb10.chm196659
f1_keywords:
- vbapb10.chm196659
ms.prod: publisher
api_name:
- Publisher.Document.Stories
ms.assetid: 4ffc7d20-eb11-942e-e28a-81c2caa19a50
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.Stories property (Publisher)

Returns a **[Stories](Publisher.Stories.md)** collection containing all stories in the publication.


## Syntax

_expression_.**Stories**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

Stories


## Example

This example assigns the first story in the **Stories** collection to a variable.

```vb
Sub FirstStory() 
 
 Dim stFirst As Story 
 
 stFirst = Application.ActiveDocument.Stories(1) 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]