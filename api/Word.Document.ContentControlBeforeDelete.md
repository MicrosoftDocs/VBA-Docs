---
title: Document.ContentControlBeforeDelete event (Word)
keywords: vbawd10.chm4001011
f1_keywords:
- vbawd10.chm4001011
ms.prod: word
api_name:
- Word.Document.ContentControlBeforeDelete
ms.assetid: a690fb97-0de3-de0e-7e84-edaaea756e83
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ContentControlBeforeDelete event (Word)

Occurs before removing a content control from a document.


## Syntax

_expression_.**ContentControlBeforeDelete'(**_OldContentControl_**, **_InUndoRedo_**)

 _expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _OldContentControl_|Required| **ContentControl**|The content control being deleted.|
| _InUndoRedo_|Required| **Boolean**| Specifies whether the removal is taking place as part an undo or redo action.|

## Remarks

For information about using events with the **Document** object, see [Using events with the Document object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]