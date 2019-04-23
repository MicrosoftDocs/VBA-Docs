---
title: Document.SelectAllEditableRanges method (Word)
keywords: vbawd10.chm158007764
f1_keywords:
- vbawd10.chm158007764
ms.prod: word
api_name:
- Word.Document.SelectAllEditableRanges
ms.assetid: 510cd397-4c39-f36b-ed59-524247b35f16
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.SelectAllEditableRanges method (Word)

Selects all ranges for which the specified user or group of users has permission to modify.


## Syntax

_expression_. `SelectAllEditableRanges`( `_EditorID_` )

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _EditorID_|Optional| **Variant**|Can be either a  **String** that represents the user's email alias (if in the same domain), an email address, or a **WdEditorType** constant that represents a group of users. If omitted, only ranges for which all users have permissions will be selected.|

## Example

The following example selects all ranges for which the current user has permission to modify.


```vb
ActiveDocument.SelectAllEditableRanges wdEditorCurrent
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]