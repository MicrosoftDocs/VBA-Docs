---
title: Document.SelectUnlinkedControls method (Word)
keywords: vbawd10.chm158007846
f1_keywords:
- vbawd10.chm158007846
ms.prod: word
api_name:
- Word.Document.SelectUnlinkedControls
ms.assetid: 6d757837-0959-6754-bfae-e840ea7de339
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.SelectUnlinkedControls method (Word)

Returns a  **[ContentControls](Word.ContentControls.md)** collection that represents all content controls in a document that are not linked to an XML node in the document's XML data store. Read-only.


## Syntax

_expression_. `SelectUnlinkedControls`( `_Stream_` )

 _expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Stream_|Optional| **CustomXMLPart**|A custom XML part reference. Setting this parameter filters the returned content controls to include only content controls that reference this  **CustomXMLPart** in their **XMLMapping** definition.|

## Return value

ContentControls


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]