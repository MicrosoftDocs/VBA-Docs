---
title: Document.ContentControlOnEnter event (Word)
keywords: vbawd10.chm4001013
f1_keywords:
- vbawd10.chm4001013
ms.prod: word
api_name:
- Word.Document.ContentControlOnEnter
ms.assetid: 593eca61-886c-85e9-fde2-1dc20c80740b
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ContentControlOnEnter event (Word)

Occurs when a user enters a content control.


## Syntax

_expression_.**ContentControlOnEnter'(**_ContentControl_**, )

 _expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ContentControl_|Required| **ContentControl**|The content control that the user is entering.|

## Remarks


> [!IMPORTANT] 
> This event fires only for the content control that you enter and not for parent content controls. For example, if you have a text box content control nested inside a group content control, and you place the cursor inside the text box content control, this event fires only once for the text box content control; it does not fire for the parent group content control.

For information about using events with the **Document** object, see [Using events with the Document object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]