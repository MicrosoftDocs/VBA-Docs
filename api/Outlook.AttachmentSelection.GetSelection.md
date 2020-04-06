---
title: AttachmentSelection.GetSelection method (Outlook)
keywords: vbaol11.chm3534
f1_keywords:
- vbaol11.chm3534
ms.prod: outlook
api_name:
- Outlook.AttachmentSelection.GetSelection
ms.assetid: 048d6d00-8928-68a5-f02c-20fdbae093c6
ms.date: 06/08/2017
localization_priority: Normal
---


# AttachmentSelection.GetSelection method (Outlook)

Returns a  **[Selection](Outlook.Selection.md)** object that contains the kind of objects specified by the _SelectionContents_ parameter, and that are currently selected in the active explorer where the parent item of the **[AttachmentSelection](Outlook.AttachmentSelection.md)** object is.


## Syntax

_expression_. `GetSelection`( `_SelectionContents_` )

_expression_ A variable that represents an [AttachmentSelection](Outlook.AttachmentSelection.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SelectionContents_|Required| **OlSelectionContents**|Specifies the kind of objects in the selection to return.|

## Return value

A  **Selection** object that contains the specified kind of objects that are selected in the active explorer.


## Remarks

The only reason that this method is exposed on the  **AttachmentSelection** object is because the **AttachmentSelection** inherits from the **Selection** object. This method is not intended to be called on the **AttachmentSelection** object.


## See also


[AttachmentSelection Object](Outlook.AttachmentSelection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]