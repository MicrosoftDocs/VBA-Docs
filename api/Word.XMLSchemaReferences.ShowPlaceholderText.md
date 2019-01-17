---
title: XMLSchemaReferences.ShowPlaceholderText property (Word)
keywords: vbawd10.chm116129799
f1_keywords:
- vbawd10.chm116129799
ms.prod: word
api_name:
- Word.XMLSchemaReferences.ShowPlaceholderText
ms.assetid: 432b25c0-79a1-7930-d0a5-c69a0e50bf72
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLSchemaReferences.ShowPlaceholderText property (Word)

Returns or sets a  **Boolean** that represents whether automatic placeholder text is displayed for XML elements in a document. Read/write.


## Syntax

 _expression_. `ShowPlaceholderText`

 _expression_ An expression that returns an '[XMLSchemaReferences](Word.XMLSchemaReferences.md)' collection.


## Remarks

 **True** displays placeholder text. **False** hides placeholder text.


## Example

The following toggles the display of placeholder text in the active document.


```vb
ActiveDocument.XMLSchemaReferences.ShowPlaceholderText = _ 
 Not ActiveDocument.XMLSchemaReferences.ShowPlaceholderText
```


## See also


[XMLSchemaReferences Collection](Word.XMLSchemaReferences.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]