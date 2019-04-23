---
title: Debug menu
keywords: vbui6.chm2057561
f1_keywords:
- vbui6.chm2057561
ms.prod: office
ms.assetid: 521e91a6-53bd-e0cc-a20c-fc82ba58b28c
ms.date: 11/24/2018
localization_priority: Normal
---


# Debug menu

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Add Watch**| Displays the **[Add Watch](add-watch-dialog-box.md)** dialog box in which you enter a [watch expression](../../Glossary/vbe-glossary.md#watch-expression). The expression can be any valid Basic expression. Watch expressions are updated in the [Watch window](watch-window.md) each time you enter [break mode](../../Glossary/vbe-glossary.md#break-mode).|![Add Watch toolbar button](../../../images/tbr_addw_ZA01201668.gif) | |
|**Clear All Breakpoints**| Removes all [breakpoints](../../Glossary/vbe-glossary.md#breakpoint) in your project. Your application may still interrupt execution, however, if you have set a watch expression or selected the **Break on All Errors** option on the **[General](general-tab-options-dialog-box.md)** tab of the **[Options](options-dialog-box.md)** dialog box. You cannot undo the **Clear All Breakpoints** command. | ![Clear All Breakpoints toolbar button](../../../images/tbr_clbp_ZA01201686.gif) |CTRL+SHIFT+F9 |
|**Compile `<project>`** | Compiles your project. | ![Compile project toolbar button](../../../images/tbr_comp_ZA01201690.gif)| |
|**Edit Watch** | Displays the **[Edit Watch](edit-watch-dialog-box.md)** dialog box in which you can edit or delete a watch expression. Available when the watch is set even if the Watch window is hidden.|![Edit Watch toolbar button](../../../images/tbr_edtw_ZA01201700.gif) | CTRL+W |
|**Run to Cursor** | When your application is in design mode, use **Run To Cursor** to select a statement further down in your code where you want execution to stop. Your application will run from the current statement to the selected statement, and the current line of execution margin indicator (![Run to cursor](../../../images/wcurline_ZA01201810.gif)) appears on the **Margin Indicator** bar.<br/><br/>You can use this command, for example, to avoid stepping through large loops. | | CTRL+F8|
|**Set Next Statement** | Sets the execution point to the line of code you choose. You can set a different line of code to execute after the currently selected statement by selecting the line of code you want to execute and choosing the **Set Next Statement** command, or by dragging the **Current Execution Line** margin indicator to the line of code that you want to execute.<br/><br/>Using **Set Next Statement**, you can choose a line of code located before or after the currently selected statement. When you run the code, any intervening code isn't executed. Use this command when you want to rerun a statement within the current procedure or to skip over statements you don't want to execute. You can't use **Set Next Statement** for statements in different procedures. | ![Set Next Statement toolbar button](../../../images/tbr_snst_ZA01201746.gif) | CTRL+F9 |
|**Show Next Statement** | Highlights the next statement to be executed. Use the **Show Next Statement** command to place the cursor on the line that will execute next. Available only in break mode.|![Show Next Statement toolbar button](../../../images/tbr_shns_ZA01201743.gif) | |
|**Step Into** | Executes code one statement at a time.<br/><br/>When not in design mode, **Step Into** enters break mode at the current line of execution. If the statement is a call to a procedure, the next statement displayed is the first statement in the procedure.<br/><br/>At [design time](../../Glossary/vbe-glossary.md#design-time), this menu item begins execution and enters break mode before the first line of code is executed. If there is no current execution point, the **Step Into** command may appear to do nothing until you do something that triggers code, for example click on a document.|![Step Into toolbar button](../../../images/tbr_stpi_ZA01201749.gif) | F8 |
|**Step Over** | Similar to **Step Into**. The difference in use occurs when the current statement contains a call to a procedure. **Step Over** executes the procedure as a unit, and then steps to the next statement in the current procedure. Therefore, the next statement displayed is the next statement in the current procedure regardless of whether the current statement is a call to another procedure. Available in break mode only.| ![Step Over toolbar button](../../../images/tbr_stpo_ZA01201750.gif) | SHIFT+F8 |
|**Step Out** |Executes the remaining lines of a function in which the current execution point lies. The next statement displayed is the statement following the procedure call. All of the code is executed between the current and the final execution points. Available in break mode only.|![Step Out toolbar button](../../../images/tbr_stot_ZA01201748.gif) |CTRL+SHIFT+F8|
|**Toggle Breakpoint** |Sets or removes a breakpoint at the current line. You can't set a breakpoint on lines containing nonexecutable code such as comments, declaration statements, or blank lines.<br/><br/>A line of code in which a breakpoint is set appears in the colors specified on the **[Editor Format](editor-format-tab-options-dialog-box.md)** tab of the **Options** dialog box.| ![Toggle Breakpoint toolbar button](../../../images/tbr_bkpt_ZA01201681.gif) | F9 |

## See also

- [Debug toolbar](debug-toolbar.md)
- [Menus and commands](../menus-commands.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
