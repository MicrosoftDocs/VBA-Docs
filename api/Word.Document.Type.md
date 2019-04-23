---
title: Document.Type property (Word)
keywords: vbawd10.chm158007306
f1_keywords:
- vbawd10.chm158007306
ms.prod: word
api_name:
- Word.Document.Type
ms.assetid: 8fcf6280-5fbc-10bf-95ef-7461c02102d2
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Type property (Word)

Returns the document type (template or document). Read-only  **[WdDocumentType](Word.WdDocumentType.md)**.


## Syntax

_expression_.**Type**

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Example

If the active window contains a document, this example redefines the Heading 1 style as centered.


```vb
If ActiveDocument.ActiveWindow.Type = wdWindowDocument Then 
 ActiveDocument.Styles("Heading 1") _ 
 .ParagraphFormat.Alignment = wdAlignParagraphCenter 
End If
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]