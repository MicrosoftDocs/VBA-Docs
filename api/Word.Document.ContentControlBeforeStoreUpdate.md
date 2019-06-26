---
title: Document.ContentControlBeforeStoreUpdate event (Word)
keywords: vbawd10.chm4001014
f1_keywords:
- vbawd10.chm4001014
ms.prod: word
api_name:
- Word.Document.ContentControlBeforeStoreUpdate
ms.assetid: a73aae31-bd03-1422-dbf2-1e7943d4a08a
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ContentControlBeforeStoreUpdate event (Word)

Occurs before updating the document's XML data store with the value of a content control.


## Syntax

_expression_.**ContentControlBeforeStoreUpdate'(**_ContentControl_**, **_Content_**)

 _expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ContentControl_|Required| **ContentControl**|The content control being updated.|
| _Content_|Required| **String**|The content being stored for a control in the document data store. Use this parameter to change the XML data before sending the value to the XML data store.|

## Remarks


> [!NOTE] 
> This event does not occur for repeating content controls.

For information about using events with the  **Document** object, see [Using events with the Document object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]