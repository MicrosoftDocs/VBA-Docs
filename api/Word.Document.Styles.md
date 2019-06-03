---
title: Document.Styles property (Word)
keywords: vbawd10.chm158007318
f1_keywords:
- vbawd10.chm158007318
ms.prod: word
api_name:
- Word.Document.Styles
ms.assetid: 30784574-92d1-a2fa-1032-6e1f8bb79ccf
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Styles property (Word)

Returns a  **[Styles](Word.styles.md)** collection for the specified document. Read-only.


## Syntax

_expression_.**Styles**

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example applies the Heading 1 style to each paragraph in the active document that begins with the word "Chapter."


```vb
For Each para In ActiveDocument.Paragraphs 
 If para.Range.Words(1).Text = "Chapter " Then 
 para.Style = ActiveDocument.Styles(wdStyleHeading1) 
 End If 
Next para
```

This example opens the template attached to the active document and modifies the Heading 1 style. The template is saved, and all styles in the active document are updated.




```vb
Set tempDoc = ActiveDocument.AttachedTemplate.OpenAsDocument 
With tempDoc.Styles(wdStyleHeading1).Font 
 .Name = "Arial" 
 .Size = 16 
End With 
tempDoc.Close SaveChanges:=wdSaveChanges 
ActiveDocument.UpdateStyles
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]