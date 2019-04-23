---
title: DoCmd.SetMenuItem method (Access)
keywords: vbaac10.chm4181
f1_keywords:
- vbaac10.chm4181
ms.prod: access
api_name:
- Access.DoCmd.SetMenuItem
ms.assetid: 690263c1-5e0f-54cd-1032-b2f718d82075
ms.date: 02/16/2019
localization_priority: Normal
---


# DoCmd.SetMenuItem method (Access)

The **SetMenuItem** method carries out the SetMenuItem action in Visual Basic.


## Syntax

_expression_.**SetMenuItem** (_MenuIndex_, _CommandIndex_, _SubcommandIndex_, _Flag_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MenuIndex_|Required|**Variant**|An integer, counting from 0, that is the valid index of a menu on the custom menu bar or global menu bar for the active window, as defined in the menu bar macro for the custom menu bar or global menu bar. <br/><br/>If you select a menu with this argument and leave the _CommandIndex_ and _SubcommandIndex_ arguments blank (or set them to 1), you can enable or disable the menu name itself. You can't, however, check or uncheck a menu name (Microsoft Access ignores the **acMenuCheck** and **acMenuUncheck** settings for the _Flag_ argument for menu names).|
| _CommandIndex_|Optional|**Variant**|An integer, counting from 0, that's the valid index of a command on the menu selected by the _MenuIndex_ argument, as defined in the macro group that defines the selected menu for the custom menu bar or global menu bar for the active window.|
| _SubcommandIndex_|Optional|**Variant**|An integer, counting from 0, that's the valid index of a subcommand in the submenu selected by the _CommandIndex_ argument, as defined in the macro group that defines the selected submenu for the custom menu bar or global menu bar for the active window.|
| _Flag_|Optional|**Variant**|The state you want to set the command or subcommand to. Can be one of the following constants:<ul><li><b>acMenuCheck</b></li><li><b>acMenuGray</b></li><li><b>acMenuUncheck</b></li><li><b>acMenuUngray</b>  (default)</li></ul>|

## Remarks

You can use the **SetMenuItem** method to set the state of menu items (enabled or disabled, checked or unchecked) on the custom menu bar or global menu bar for the active window.

> [!NOTE] 
> The **SetMenuItem** method works only with custom menu bars and global menu bars created by using menu bar macros. The **SetMenuItem** method is included in this version of Access only for compatibility with versions prior to Access 97. It doesn't work with the new command bars functionality.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]