---
title: Immediate window
keywords: vblr6.chm2058895
f1_keywords:
- vblr6.chm2058895
ms.prod: office
ms.assetid: e2e16178-0216-d91f-5c59-bd39574be84a
ms.date: 11/21/2018
localization_priority: Normal
---


# Immediate window

![Immediate window](../../../images/immed_ZA01201615.gif)

Allows you to:

- Type or paste a line of code and press ENTER to run it.
    
- Copy and paste the code from the Immediate window into the Code window, but does not allow you to save code in the Immediate window.
    
The Immediate window can be dragged and positioned anywhere on your screen unless you have made it a dockable window from the **Docking** tab of the **[Options](options-dialog-box.md)** dialog box.

You can close the window by selecting the **Close** box. If the **Close** box is not visible, double-click the title bar to make the **Close** box visible, and then select it.

> [!NOTE] 
> In break mode, a statement in the Immediate window is executed in the context or [scope](../../Glossary/vbe-glossary.md#scope) that is displayed in the **Procedure** box. For example, if you type **Print**_variablename_, your output is the value of a local variable. This is the same as if the **Print** method had occurred in the procedure you were executing when the program halted.

## Keyboard shortcuts

Use these key combinations in the Immediate window.

|Press|To|
|:-----|:-----|
|ENTER|Run a line of selected code.|
|CTRL+C|Copy the selected text to the **Clipboard**.|
|CTRL+V|Paste the **Clipboard** contents at the insertion point.|
|CTRL+X|Cut the selected text to the **Clipboard**.|
|CTRL+L|Display **[Call Stack](call-stack-dialog-box.md)** dialog box (break mode only).|
|F5|Continue running an application.|
|F8|Execute code one line at a time (single step).|
|SHIFT+F8|Execute code one procedure at a time (procedure step).|
|DELETE or DEL|Delete the selected text without placing it on the **Clipboard**.|
|F2|Display the **[Object Browser](object-browser.md)**.|
|CTRL+ENTER|Insert carriage return.|
|CTRL+HOME|Move the cursor to the top of the Immediate window.|
|CTRL+END|Move the cursor to the end of the Immediate window.|
|SHIFT+F10|View shortcut menu.|
|ALT+F5|Runs the error handler code or returns the error to the calling procedure. Does not affect the setting for error trapping on the **General** tab of the **Options** dialog box.|
|ALT+F8|Steps into the error handler or returns the error to the calling procedure. Does not affect the setting for error trapping on the **General** tab of the **Options** dialog box.|

## See also

- [Use the Immediate window](use-the-immediate-window.md)
- [Window elements](../window-elements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]