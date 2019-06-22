---
title: Document.New event (Word)
keywords: vbawd10.chm4001004
f1_keywords:
- vbawd10.chm4001004
ms.prod: word
api_name:
- Word.Document.New
ms.assetid: c37f7e20-f429-e921-3d17-609d536e8baa
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.New event (Word)

Occurs when a new document based on the template is created. A procedure for the  **New** event will run only if it is stored in a template.


## Syntax

_expression_.**Private Sub Document_New'

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about using events with the  **Document** object, see [Using events with the Document object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).


## Example

This example asks the user whether to save all other open documents when a new document based on the template is created. (This procedure is stored in the  **ThisDocument** class module of a template, not a document.)


```vb
Private Sub Document_New() 
 Dim intResponse As Integer 
 Dim strName As String 
 Dim docLoop As Document 
 
 intResponse = MsgBox("Save all other documents?", vbYesNo) 
 
 If intResponse = vbYes Then 
 strName = ActiveDocument.Name 
 For Each docLoop In Application.Documents 
 With docLoop 
 If .Name <> strName Then 
 .Save 
 End If 
 End With 
 Next docLoop 
 End If 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]