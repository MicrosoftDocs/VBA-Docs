---
title: Document.ContentControlAfterAdd event (Word)
keywords: vbawd10.chm4001010
f1_keywords:
- vbawd10.chm4001010
ms.prod: word
api_name:
- Word.Document.ContentControlAfterAdd
ms.assetid: 9a19d147-76bd-eb92-5844-c56b2d6eae7c
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ContentControlAfterAdd event (Word)

Occurs after adding a content control to a document.


## Syntax

_expression_.**ContentControlAfterAdd'(**_NewContentControl_**, **_InUndoRedo_**)

 _expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NewContentControl_|Required| **ContentControl**|The content control being added.|
| _InUndoRedo_|Required| **Boolean**|Specifies whether the addition is taking place as part an undo or redo action.|

## Remarks

For information about using events with the **Document** object, see [Using events with the Document object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]