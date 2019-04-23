---
title: Explorer.RemoveFromSelection method (Outlook)
keywords: vbaol11.chm3310
f1_keywords:
- vbaol11.chm3310
ms.prod: outlook
api_name:
- Outlook.Explorer.RemoveFromSelection
ms.assetid: f31bc78f-500e-2f73-ea14-8d5f19cd44e9
ms.date: 02/16/2019
localization_priority: Normal
---


# Explorer.RemoveFromSelection method (Outlook)

Cancels the selection of the specified Microsoft Outlook item in the active explorer.


## Syntax

_expression_.**RemoveFromSelection** (_Item_)

_expression_ A variable that represents an **[Explorer](Outlook.Explorer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item to be removed from the selection.|

## Remarks

The selection in the active explorer is represented by the **[Selection](Outlook.Selection.md)** object returned by the **[Explorer.Selection](Outlook.Explorer.Selection.md)** property.

To be removed from a selection, an item must be selectable in the current view of the active explorer. However, the item does not have to be visible in the view.

Outlook will return an error when you call the **RemoveFromSelection** method under the following conditions:

- The specified item is not in the current view of the active explorer.    
- The specified item is being edited in the current view of the active explorer.    
- The current view has been filtered, and the application of the filter removed the item from the view.   
- The specified item has not been saved.   
- The specified item represents a **[StorageItem](Outlook.StorageItem.md)**.   
- The current view is a conversation view.    
- No current view exists for the active explorer.
    
If the specified item is selected, calling **RemoveFromSelection** will cause the **[SelectionChange](Outlook.Explorer.SelectionChange.md)** event to fire. If the item is not selected, calling **RemoveFromSelection** will not cause the **SelectionChange** event to fire.

Calling **RemoveFromSelection** does not scroll the view to make the specified item visible in the view and does not expand or collapse groups in the view.

The following table illustrates the results of calling **RemoveFromSelection**, taking into consideration any current selection (the **[Selection.Count](Outlook.Selection.Count.md)** property), whether the Reading pane is displayed, and whether the specified item is displayed in the Reading pane.

|Existing Selection.Count|Reading pane displayed|Specified item displayed in Reading pane|Results|
|:-----------------------|:---------------------|:---------------------------------------|:------|
|1|Yes|Yes|<ul><li>The selection is cleared.</li><li><b>SelectionChange</b> fires.</li><li>Reading pane is empty.</li></ul>|
|>1|Yes|No|<ul><li>The item is removed from the selection.</li><li><b>SelectionChange</b> fires.</li><li>Reading pane does not change.</li></ul>|
|>1|Yes|Yes|<ul><li>The item is removed from the selection.</li><li><b>SelectionChange</b> fires.</li><li>Reading pane displays the next item or adjacent item in the selection.</li></ul>|
|>=1|No|N/A|<ul><li>The item is removed from the selection.</li><li><b>SelectionChange</b> fires.</li></ul>|

If the specified item exists in the current view but is not selected in that view, calling **RemoveFromSelection** does not result in any change to the selection and does not fire the **SelectionChange** event.

When you specify an item in a recurring appointment or task as an argument to the **RemoveFromSelection** method, make sure that before you pass the argument, you obtain an instance of the occurrence by first expanding the recurrences by using the **[IncludeRecurrences](Outlook.Items.IncludeRecurrences.md)** property and the **[Items](Outlook.Items.md)** collection. If you do not expand the recurrences and obtain an occurrence in the series, you would be passing an instance variable that represents the appointment or task series, and the **RemoveFromSelection** method would be operating on the series instead of the occurrence.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]