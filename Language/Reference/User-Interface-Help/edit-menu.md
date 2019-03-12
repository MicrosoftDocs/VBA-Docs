---
title: Edit menu
keywords: vbui6.chm2057557
f1_keywords:
- vbui6.chm2057557
ms.prod: office
ms.assetid: 1f84bfae-a6a2-34f4-67c5-c50c7dab8b73
ms.date: 11/24/2018
localization_priority: Normal
---


# Edit menu

## Bookmarks

Displays a menu that you can use to create or remove placeholders in the [Code window](code-window.md), move to the next or preceding bookmark, or clear all of the bookmarks. Bookmarks mark lines of code so that you can easily return to them at a later time. When you add a bookmark, a ![Bookmark](../../../images/wbkmark_ZA01201807.gif) appears next to the line of code where the bookmark is inserted.

|Command|Description|Toolbar button|
|:------|:----------|:-------------|
|**Toggle Bookmark** | Toggles a bookmark on or off.|![Toggle Bookmark toolbar button](../../../images/tbr_tbmk_ZA01201753.gif)|
|**Next Bookmark** | Moves the insertion point to the next bookmark. |![Next Bookmark toolbar button](../../../images/tbr_nxtb_ZA01201717.gif)| 
|**Previous Bookmark** | Moves the insertion point to the previous bookmark.|![Previous Bookmark toolbar button](../../../images/tbr_prvb_ZA01201729.gif)| 
|**Clear All Bookmarks**| Removes all bookmarks.|![Clear All Bookmarks toolbar button](../../../images/tbr_clrb_ZA01201687.gif) | 

## Clear, Delete

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Clear** |Deletes text only when a module is active.| | |
|**Delete** |At other times, changes to the **Delete** command that deletes the currently selected control, text, or [watch expression](../../Glossary/vbe-glossary.md#watch-expression). You can undo the **Delete** command in the Code window or in the form if you deleted the control from the form. Not available at [run time](../../Glossary/vbe-glossary.md#run-time). |![Delete Toolbar button](../../../images/tbr_del_ZA01201696.gif) | DELETE |


## Comment Block, Uncomment Block

Adds and removes the comment character, an apostrophe, for each line of a selected block of text. If you do not have text selected and you choose the **Comment Block** or **Uncomment Block** command, the comment character is added or removed in the line where the pointer is located.

|Command|Description|Toolbar button|
|:------|:----------|:-------------|
|**Comment Block**|Adds the comment character to each line of a selected block of text.|![Comment Block Toolbar button](../../../images/tbr_comt_ZA01201691.gif)|
|**Uncomment Block**|Removes the comment character from each line of a selected block of text.|![Uncomment Block Toolbar button](../../../images/tbr_uncm_ZA01201761.gif)|


## Complete Word

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Complete Word** | Fills in the rest of the word you are typing after you have entered enough characters for Visual Basic to identify the word you want. |![Complete Word Toolbar button](../../../images/tbr_cwrd_ZA01201695.gif) | CTRL+SPACEBAR |


## Cut, Copy, Paste, Delete

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Cut** |Removes the selected control or text and places it on the **Clipboard**. You must select at least one character or control for this command to be available. You can undo the **Cut** command when used on a control.|![Cut Toolbar button](../../../images/tbr_cut_ZA01201694.gif) |CTRL+X or SHIFT+DELETE |
|**Copy** |Copies the selected control or text onto the **Clipboard**. You must select at least one character or control for this command to be available. You cannot undo the **Copy** command in the Code window.|![Copy Toolbar button](../../../images/tbr_copy_ZA01201692.gif) |CTRL+C or CTRL+INSERT |
|**Paste** |Inserts the contents of the **Clipboard** at the current location. Text is placed at the insertion point. Pasted controls are placed in the middle of the form. You can undo the **Paste** command in the Code window or in the form if you pasted the control onto the form.|![Paste Toolbar button](../../../images/tbr_pste_ZA01201730.gif) |CTRL+V or SHIFT+INS |
|**Delete** |Deletes the currently selected control, text, or [watch expression](../../Glossary/vbe-glossary.md#watch-expression). You can undo the **Delete** command only in the Code window. Not available at [run time](../../Glossary/vbe-glossary.md#run-time).<br/><br/>**NOTE**: To delete a file from your disk, use the standard deletion procedures for your operating system.|![Delete Toolbar button](../../../images/tbr_del_ZA01201696.gif) |DEL |


## Find, Find Next

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Find** |Searches for the specified text in a search range specified in the **[Find](find-dialog-box.md)** dialog box. If a search is successful, the **Find** dialog box closes and Visual Basic selects the located text. If no match is found, Visual Basic displays a message stating that the text was not found.|![Find Toolbar button](../../../images/tbr_find_ZA01201703.gif) | CTRL+F |
|**Find Next** |Finds and selects the next occurrence of the text specified in the **Find What** box of the **Find** dialog box.|![Find Next Toolbar button](../../../images/tbr_next_ZA01201716.gif) |SHIFT+F4 (**Find Next**) or SHIFT+F3 (**Find Previous**) |


## Indent, Outdent

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Indent**|Shifts all lines in the selection to the next tab stop. If you place the cursor anywhere in a line and choose the **Indent** command, the entire line is shifted to the next tab stop. All lines in the selection are moved the same number of spaces to retain the same relative indentation within the selected block.<br/><br/>You can change the tab width on the **[Editor](editor-tab-options-dialog-box.md)** tab of the **[Options](options-dialog-box.md)** dialog box.|![Indent Toolbar button](../../../images/tbr_inde_ZA01201711.gif)|CTRL+M |
|**Outdent**|Shifts all lines in the selection to the previous tab stop. If you place the cursor anywhere in a line and choose the **Outdent** command, the entire line is shifted to the previous tab stop. All lines in the selection are moved the same number of spaces to retain the same relative indentation within the selected block.<br/><br/>You can change the tab width on the **Editor** tab of the **Options** dialog box.|![Outdent Toolbar button](../../../images/tbr_outd_ZA01201721.gif) | CTRL+SHIFT+M |

## Parameter Info, Quick Info

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Parameter Info** |Shows a popup in the Code window that contains information about the parameters of the initial function or statement. If you have a function or statement that contains functions as its parameters, choosing **Parameter Info** provides information about the first function. **Quick Info** provides information about each embedded function.<br/><br/>As you type a parameter, it is bold until you type the comma used to delineate it from the next parameter.<br/><br/>The **Parameter Info**, once activated, will not close until:<br/>- All of the required parameters are entered.<br/>- The function is ended without using all of the optional parameters.<br/>- You press ESC.|![Parameter Info Toolbar button](../../../images/tbr_ptip_ZA01201731.gif) |CTRL+SHIFT+I |
|**Quick Info**|Provides the syntax for a variable, function, statement, method, or procedure selected in the Code window. **Quick Info** shows the syntax for the item and highlights the current parameter. For functions and procedures with parameters, the parameter appears bold as you type it until you type the comma used to delineate it from the next parameter.<br/><br/>To have **Quick Info** automatically appear as you type your code, select **Auto Quick Info** on the **Editor** tab in the **Options** dialog box.|![Quick Info Toolbar button](../../../images/tbr_qtip_ZA01201732.gif) | CTRL+I |

## Replace

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Replace**| Searches code in the project for the specified text and replaces it with the new text specified in the **[Replace](replace-dialog-box.md)** dialog box.|![Replace Toolbar button](../../../images/tbr_repl_ZA01201735.gif) |CTRL+H |


## Select All, Select Constant, Select Member

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Select All**|Selects all of the code in the active Code window or all the controls on a form.| | |
|**Select Constant** |Opens a drop-down list box in the Code window that contains the valid constants for a property that you typed, and that preceded the equal sign (=). The **List Constants** command also works for functions with arguments that are constants. To have the list box automatically open as you type your code, select **Auto List Members** on the **Editor** tab in the **Options** dialog box.<br/><br/>You can find the constant you want by:<br/>- Typing the name.<br/>- Using the up and down arrow keys to move up and down in the list.<br/>- Scrolling through the list and selecting the constant you want.<br/><br/>You can insert the constant into your code statement by:<br/>- Double-clicking the constant.<br/>- Selecting the constant and pressing TAB to insert the selection or pressing ENTER to insert the selection and move to the next line.|![Select Constant Toolbar button](../../../images/tbr_selc_ZA01201740.gif) | CTRL+SHIFT+J |
|**Select Member** |Opens a drop-down list box in the Code window that contains the properties and methods available for the object. The **List Properties/Methods** command also displays a list of the globally available methods when the pointer is on a blank space. To have the list box automatically open as you type your code, select **Auto List Members** on the **Editor** tab in the **Options** dialog box.<br/><br/>You can find the property or method you want in the list box by:<br/>- Typing the name. As you type, the property or method that matches the characters you type is selected and moves to the top of the list.<br/>- Using the up and down arrow keys to move up and down in the list.<br/>- Scrolling through the list and selecting the property or method you want.<br/><br/>You can insert the property or method into your statement by:<br/>- Double-clicking the property or method.<br/>- Selecting the property or method and pressing TAB to insert the selection or pressing ENTER to insert the selection and move to the next line.<br/><br/>**NOTE**: Objects of the type **Variant** do not show a list after the period (`.`). |![Select Member Toolbar button](../../../images/tbr_selm_ZA01201741.gif) | CTRL+J |

## Undo, Redo

For text edits, you can use **Undo** and **Redo** to restore up to twenty edits. These commands are unavailable at runtime, or if there was no previous edit, or if any other action has been performed after the last edit. Also, some large edits may cause low memory conditions that could prevent an **Undo** action.

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Undo** |Reverses the last editing action, such as typing text in the Code window or deleting controls. When you delete one or more controls, you can use the **Undo** command to restore the controls and all their properties.<br/><br/>**NOTE**: You can't undo a **Cut** operation by using the **Undo** command on a form.|![Undo Toolbar button](../../../images/tbr_undo_ZA01201762.gif) | CTRL+Z or ALT+BACKSPACE |
|**Redo** |Restores the last text editing or resizing and positioning of controls if no other actions have occurred since the last **Undo**. |![Redo Toolbar button](../../../images/tbr_redo_ZA01201734.gif) | |

## See also

- [Editor toolbar](editor-toolbar.md)
- [Menus and commands](../menus-commands.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
