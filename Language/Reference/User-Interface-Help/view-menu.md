---
title: View menu
keywords: vbui6.chm2057558
f1_keywords:
- vbui6.chm2057558
ms.prod: office
ms.assetid: 1c6bc77b-0a89-6c5d-eec7-30bb29ad67c9
ms.date: 11/24/2018
localization_priority: Normal
---


# View menu

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Call Stack** |Displays the **[Call Stack](call-stack-dialog-box.md)** dialog box, which lists the procedures that have started but are not completed. Available only in [break mode](../../Glossary/vbe-glossary.md#break-mode).<br/><br/>When Visual Basic is executing the code in a procedure, that procedure is added to a list of active procedure calls. If that procedure then calls another procedure, there are two procedures on the list of active procedure calls. Each time a procedure calls another [Sub](../../Glossary/vbe-glossary.md#sub-procedure), [Function](../../Glossary/vbe-glossary.md#function-procedure), or [Property](../../Glossary/vbe-glossary.md#property-procedure) procedure, it is added to the list. Each procedure is removed from the list as execution is returned to the calling procedure. Procedures called from the [Immediate window](immediate-window.md) are also added to the calls list.<br/><br/>You can also display the **Call Stack** dialog box by clicking the **Calls** button (...) next to the **Procedure** box in the [Locals window](locals-window.md).|![Call Stack Toolbar button](../../../images/tbr_call_ZA01201683.gif) |CTRL+L|
|**Code** |Displays or activates the [Code window](code-window.md) for a currently selected [object](../../Glossary/vbe-glossary.md#object).|![Code Toolbar button](../../../images/tbr_code_ZA01201689.gif) |F7|
|**Definition** |Displays the location in the Code window where the variable or procedure under the pointer is defined. If the definition is in a referenced library, it is displayed in the **Object Browser**.| |SHIFT+F2|
|**`<Host application>`** |Moves the host application on top of the Visual Basic Editor so that you can view it. The name of the command changes to the name of the host application.| |ALT+F11|
|**Immediate Window** |Displays the [Immediate window](immediate-window.md), and displays information resulting from debugging statements in your code or from commands typed directly into the window.<br/><br/>Use the Immediate window to:<br/>- Test problematic or newly written code.<br/>- Query or change the value of a variable while running an application. While execution is halted, assign the variable a new value as you would in code.<br/>- Query or change a property value while running an application.<br/>- Call procedures as you would in code.<br/>- View debugging output while the program is running.|![Immediate Window Toolbar button](../../../images/tbr_imwd_ZA01201710.gif) |CTRL+G|
|**Last Position** |Allows you to quickly navigate to a previous location in your code. Enabled only if you edited code or made a **Definition** command call and only when the Code window is displayed. Visual Basic only keeps track of the last 8 lines that were accessed or edited.| |CTRL+SHIFT+F2|
|**Locals Window** |Displays the [Locals window](locals-window.md) and automatically displays all of the [variables](../../Glossary/vbe-glossary.md#variable) in the current stack and their values. The Locals window is automatically updated every time you change from [run time](../../Glossary/vbe-glossary.md#run-time) to [break mode](../../Glossary/vbe-glossary.md#break-mode) and every time the stack context changes.|![Locals Window Toolbar button](../../../images/tbr_lowd_ZA01201713.gif) | |
|**Object Browser** |Displays the [Object Browser](../../Glossary/vbe-glossary.md#object-browser), which lists the [object libraries](../../Glossary/vbe-glossary.md#object-library), the [type libraries](../../Glossary/vbe-glossary.md#type-library), [classes](../../Glossary/vbe-glossary.md#class), methods, properties, events, and constants you can use in code, as well as the modules and procedures you defined for your project.|![Object Browser Toolbar button](../../../images/tbr_obbr_ZA01201718.gif) | F2|
|**Object** |Displays the active item.|![Object Toolbar button](../../../images/tbr_obj_ZA01201719.gif) |SHIFT+F7|
|**Project Explorer** |Displays the **Project Explorer**, which displays a hierarchical list of the currently open projects and their contents. The **Project Explorer** is a navigational and management tool only. You cannot build an application from the **Project Explorer**.|![Project Explorer Toolbar button](../../../images/tbr_pexp_ZA01201722.gif) |CTRL+R|
|**Properties Window** |Displays the [Properties window](properties-window.md), which lists the design-time properties for a selected [form](../../Glossary/vbe-glossary.md#form), [control](../../Glossary/vbe-glossary.md#control), [class](../../Glossary/vbe-glossary.md#class), [project](../../Glossary/vbe-glossary.md#project) or [module](../../Glossary/vbe-glossary.md#module).|![Properties Window Toolbar button](../../../images/tbr_prop_ZA01201727.gif) | F6|
|**Tab Order** |Displays the **[Tab Order](tab-order-dialog-box.md)** dialog box for the active **Form**.|![Tab Order Toolbar button](../../../images/tbr_tbod_ZA01201754.gif)| |
|**Toolbars** |Lists the toolbars that are built into Visual Basic and the **Customize** command. You can toggle the toolbars on and off, or drag the toolbars to different locations on you desktop.<br/><br/>**Debug**: Displays the **Debug** toolbar, which contains buttons for common debugging tasks.<br/>**Editor**: Displays the **Editor** toolbar, which contains buttons for common editing tasks.<br/>**Standard**: Displays the **Standard** toolbar, which is the default toolbar.<br/>**UserForm**: Displays the **UserForm** toolbar, which contains buttons specific to the form.<br/>**Customize**: Displays the **[Customize](customize-dialog-box.md)** dialog box, where you can customize or create toolbars and your menu bar.| | |
|**Toolbox** |Displays or hides the **[Toolbox](toolbox.md)**, which contains the controls currently available to your application.|![Toolbox Toolbar button](../../../images/tbr_tbx_ZA01201755.gif)| |
|**Watch Window** |Displays the [Watch window](watch-window.md) and the current watch expressions. The Watch window appears automatically if watch expressions are defined in the project. If the context of the expression isn't in [scope](../../Glossary/vbe-glossary.md#scope) when going to [break mode](../../Glossary/vbe-glossary.md#break-mode), the current value isn't displayed.|![Watch Window Toolbar button](../../../images/tbr_wawd_ZA01201768.gif)| |


## See also

- [Menus and commands](../menus-commands.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]