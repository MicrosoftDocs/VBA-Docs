---
title: TextRange.Find property (Publisher)
keywords: vbapb10.chm5308497
f1_keywords:
- vbapb10.chm5308497
ms.prod: publisher
api_name:
- Publisher.TextRange.Find
ms.assetid: 453e1507-a02d-a91b-730b-fb4a13396dbc
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.Find property (Publisher)

Returns a **[FindReplace](publisher.findreplace.md)** object from the specified **TextRange** object. The **FindReplace** object is used to perform a text search and replace in the specified text range.

## Syntax

_expression_.**Find**

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Example

The following example sets an object variable to the **FindReplace** object of the text range of the first shape in the active document. A search operation is executed that applies bold formatting to every occurrence of the word Urgent in the text range.

```vb
Dim objFind as FindReplace 
Dim fFound as Boolean 
 
Set objFind = ActiveDocument.Pages(1) _ 
 .Shapes(1).TextFrame.TextRange.Find 
fFound = True 
 
With objFind 
 .Clear 
 .FindText = "Urgent" 
 Do While fFound = True 
 fFound = .Execute 
 If Not .FoundTextRange Is Nothing Then 
 .FoundTextRange.Font.Bold = True 
 End If 
 Loop 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]