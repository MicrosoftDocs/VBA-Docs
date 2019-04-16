---
title: Selection.Location property (Outlook)
keywords: vbaol11.chm3481
f1_keywords:
- vbaol11.chm3481
ms.prod: outlook
api_name:
- Outlook.Selection.Location
ms.assetid: 8a2db72a-8db0-840e-349e-5d9d22f3affb
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Location property (Outlook)

Returns an  **[OlSelectionLocation](Outlook.OlSelectionLocation.md)** constant that specifies where in the Microsoft Outlook user interface the current selection is. Read-only


## Syntax

_expression_.**Location** 

_expression_ A variable that represents a '[Selection](Outlook.Selection.md)' object.


## Remarks

A  **Location** property with the value **olViewList** means that the current selection is in a list of items in an explorer. Calling **[Selection.GetSelection](Outlook.Selection.GetSelection.md)** with **olConversationHeaders** as the argument returns a **Selection** object with **[Selection.Count](Outlook.Selection.Count.md)** equal to the number of conversation headers in the current selection.

If the  **Location** property is not equal to **olViewList**, calling **GetSelection** with **olConversationHeaders** as the argument returns a **Selection** object with **Selection.Count** equal to 0.


## See also


[Selection Object](Outlook.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]