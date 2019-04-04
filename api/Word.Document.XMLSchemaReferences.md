---
title: Document.XMLSchemaReferences property (Word)
keywords: vbawd10.chm158007757
f1_keywords:
- vbawd10.chm158007757
ms.prod: word
api_name:
- Word.Document.XMLSchemaReferences
ms.assetid: 7008fb35-017d-2f14-0627-9b524138137c
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.XMLSchemaReferences property (Word)

Returns an XMLSchemaReferences collection that represents the schemas attached to a document.


## Syntax

_expression_. `XMLSchemaReferences`

 _expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Example

The following example reloads the first schema attached to the active document.


```vb
Dim objSchema As XMLSchemaReference 
 
Set objSchema = ActiveDocument.XMLSchemaReferences.Item(1) 
 
objSchema.Reload
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]