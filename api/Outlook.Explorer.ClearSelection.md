---
title: Explorer.ClearSelection method (Outlook)
keywords: vbaol11.chm3312
f1_keywords:
- vbaol11.chm3312
ms.prod: outlook
api_name:
- Outlook.Explorer.ClearSelection
ms.assetid: 2809b5fb-961e-fb2a-a74d-fffa4484c838
ms.date: 06/08/2017
localization_priority: Normal
---


# Explorer.ClearSelection method (Outlook)

Cancels any selection in the active explorer.


## Syntax

_expression_. `ClearSelection`

_expression_ A variable that represents an '[Explorer](Outlook.Explorer.md)' object.


## Remarks

After the  **ClearSelection** method is called, the **[Count](Outlook.Selection.Count.md)** property of the **[Selection](Outlook.Selection.md)** object that the **[Explorer.Selection](Outlook.Explorer.Selection.md)** property returns is zero. Then, the **[SelectionChange](Outlook.Explorer.SelectionChange.md)** event fires unless prior to calling of **ClearSelection**, the current view did not contain any items, the current folder was empty, or the **Count** property was already zero.

If the Reading Pane is visible and the current view is a table view, calling  **ClearSelection** renders the Reading Pane blank.

If the current view or current folder does not contain any items, calling  **ClearSelection** does not result in any change to the selection and does not fire the **SelectionChange** event.

 **ClearSelection** returns an error if the item is being edited in the current view.


## See also


[Explorer Object](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]