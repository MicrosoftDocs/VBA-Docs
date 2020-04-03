---
title: Selection.GetSelection method (Outlook)
keywords: vbaol11.chm3533
f1_keywords:
- vbaol11.chm3533
ms.prod: outlook
api_name:
- Outlook.Selection.GetSelection
ms.assetid: c6af6665-d97d-3833-1014-5b43282bafc2
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.GetSelection method (Outlook)

Returns a  **[Selection](Outlook.Selection.md)** object that contains the kind of objects specified by the _SelectionContents_ parameter, and that are currently selected in the active explorer.


## Syntax

_expression_. `GetSelection`( `_SelectionContents_` )

_expression_ A variable that represents a [Selection](Outlook.Selection.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SelectionContents_|Required| **[OlSelectionContents](Outlook.OlSelectionContents.md)**|Specifies the kind of objects in the selection to return.|

## Return value

A  **Selection** object that contains the specified kind of objects that are selected in the active explorer.


## Remarks

Calling  **GetSelection** with **olConversationHeaders** as the argument returns a **Selection** object that has the **[Location](Outlook.Selection.Location.md)** property equal to **OlSelectionLocation.olViewList**.

If the current view is not a conversation view, or, if  **Selection.Location** is not equal to **OlSelectionLocation.olViewList**, calling **GetSelection** with **olConversationHeaders** as the argument returns a **Selection** object with **[Selection.Count](Outlook.Selection.Count.md)** equal to 0.


## See also


[Selection Object](Outlook.Selection.md)




[How to: Obtain and Enumerate Selected Conversations](../outlook/Concepts/Categories-and-Conversations/obtain-and-enumerate-selected-conversations.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]