---
title: Watch window
ms.prod: office
ms.assetid: 618fc1d3-dcab-6240-ff9c-048830212663
ms.date: 11/21/2018
localization_priority: Normal
---


# Watch window

![Watch window](../../../images/watch_ZA01201806.gif)

Appears automatically when watch expressions are defined in the project.

You can:

- Change the size of the column headers by dragging its border to the right to make it larger or to the left to make it smaller.
    
- Drag a selected variable to the [Immediate window](immediate-window.md) or the Watch window.
    
- Close the window by clicking the **Close** box. If the **Close** box is not visible, double-click the title bar to make the **Close** box visible, and then select it.
    
## Window elements

|Element|Description|
|:------|:----------|
|**Expression**|Lists the watch expression with the **Watch** icon ![Toolbar button](../../../images/tbr_wawd_ZA01201768.gif) on the left.|
|**Value**|Lists the value of the expression at the time of the transition to break mode.<br/><br/>You can edit a value and then press ENTER, the UP ARROW key, the DOWN ARROW key, TAB, SHIFT+TAB, or click somewhere on the screen to validate the change. If the value is illegal, the Edit field remains active and the value is highlighted. A message box describing the error also appears. Cancel a change by pressing ESC.|
|**Type**|Lists the expression type.|
|**Context**|Lists the context of the watch expression.<br/><br/>If the context of the expression isn't in [scope](../../Glossary/vbe-glossary.md#scope) when going to break mode, the current value isn't displayed.|

## Keyboard shortcuts

Use these key combinations in the Watch window.

|Press|To|
|:-----|:-----|
|SHIFT+ENTER|Display the selected watch expression.|
|CTRL+W|Display the **[Edit Watch](edit-watch-dialog-box.md)** dialog box.|
|ENTER|Expands or collapses the selected watch value if it has a plus (+) or minus (-) sign to the left of it.|
|F2|Display the [Object Browser](object-browser.md).|
|SHIFT+F10|View shortcut menu.|


## See also

- [Add a watch expression](../../concepts/forms/add-a-watch-expression.md)
- [Delete a watch expression](../../how-to/delete-a-watch-expression.md)
- [Edit a watch expression](../../how-to/edit-a-watch-expression.md)
- [Window elements](../window-elements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]