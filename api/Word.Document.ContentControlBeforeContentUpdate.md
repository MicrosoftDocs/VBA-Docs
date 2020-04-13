---
title: Document.ContentControlBeforeContentUpdate event (Word)
keywords: vbawd10.chm4001015
f1_keywords:
- vbawd10.chm4001015
ms.prod: word
api_name:
- Word.Document.ContentControlBeforeContentUpdate
ms.assetid: 297241e3-fda9-1947-8b09-9dca97930dcf
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ContentControlBeforeContentUpdate event (Word)

Occurs before updating the content in a content control, but only when the content comes from the Office XML data store.


## Syntax

_expression_.**ContentControlBeforeContentUpdate'(**_ContentControl_**, **_Content_**)

 _expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ContentControl_|Required| **ContentControl**|The content control being updated.|
| _Content_|Required| **String**|The updated content for a control. Use this parameter to change the contents of the XML data and format it for display.|

## Remarks


> [!NOTE] 
> This event does not occur for repeating content controls.

For information about using events with the **Document** object, see [Using events with the Document object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]