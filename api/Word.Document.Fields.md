---
title: Document.Fields property (Word)
keywords: vbawd10.chm158007316
f1_keywords:
- vbawd10.chm158007316
ms.prod: word
api_name:
- Word.Document.Fields
ms.assetid: 78707979-5d25-0168-2dba-ce88a2b26f9d
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Fields property (Word)

Returns a  **[Fields](Word.fields.md)** collection that represents all the fields in the document. Read-only.


## Syntax

_expression_. `Fields`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example updates all the fields in the active document.


```vb
ActiveDocument.Fields.Update
```

This example removes all the fields from the main text story and the footer in the active document.




```vb
For Each aField in ActiveDocument.Fields 
 aField.Delete 
Next aField 
Set myRange = ActiveDocument.Sections(1).Footers _ 
 (wdHeaderFooterPrimary).Range 
For Each aField In myRange.Fields 
 aField.Delete 
Next aField
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]