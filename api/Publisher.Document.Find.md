---
title: Document.Find property (Publisher)
keywords: vbapb10.chm196725
f1_keywords:
- vbapb10.chm196725
ms.prod: publisher
api_name:
- Publisher.Document.Find
ms.assetid: e9b31937-4504-79b5-5913-b2ef0a23f2a7
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.Find property (Publisher)

Returns a **[FindReplace](publisher.findreplace.md)** object from the specified **Document** object. The **FindReplace** object is used to perform a text search and replace in the specified document.

## Syntax

_expression_.**Find**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Example

Applies to the **Document** object. The following example sets an object variable to the **FindReplace** object of the active document. A search operation is executed that applies bold formatting to every occurrence of the word Important.

```vb
Dim objFind as FindReplace 
Dim fFound as Boolean 
 
Set objFind = ActiveDocument.Find 
fFound = True 
 
With objFind 
 .Clear 
 .FindText = "Important" 
 Do While fFound = True 
 fFound = .Execute 
 If Not .FoundTextRange Is Nothing Then 
 .FoundTextRange.Font.Bold = True 
 End If 
 Loop 
End With 
```

<br/>

Applies to the **[TextRange](publisher.textrange.md)** object. The following example sets an object variable to the **FindReplace** object of the text range of the first shape in the active document. A search operation is executed that applies bold formatting to every occurrence of the word Urgent in the text range.

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