---
title: Document.ContentControlBeforeStoreUpdate Event (Word)
keywords: vbawd10.chm4001014
f1_keywords:
- vbawd10.chm4001014
ms.prod: word
api_name:
- Word.Document.ContentControlBeforeStoreUpdate
ms.assetid: a73aae31-bd03-1422-dbf2-1e7943d4a08a
ms.date: 06/08/2017
---


# Document.ContentControlBeforeStoreUpdate Event (Word)

Occurs before updating the document's XML data store with the value of a content control.


## Syntax

Private Sub  _expression_ _'ContentControlBeforeStoreUpdate'(**_ContentControl_** , **_Content_**)

 _expression_ An expression that returns a '[Document](Word.Document.md)' object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ContentControl_|Required| **ContentControl**|The content control being updated.|
| _Content_|Required| **String**|The content being stored for a control in the document data store. Use this parameter to change the XML data before sending the value to the XML data store.|

## Remarks


 **Note**  This event does not occur for repeating content controls.

For information about using events with the  **Document** object, see [Using Events with the Document Object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).


## See also


[Document Object](Word.Document.md)

