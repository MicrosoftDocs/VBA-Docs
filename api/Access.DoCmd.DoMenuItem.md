---
title: DoCmd.DoMenuItem method (Access)
keywords: vbaac10.chm4148
f1_keywords:
- vbaac10.chm4148
ms.prod: access
api_name:
- Access.DoCmd.DoMenuItem
ms.assetid: b897bfdb-7f03-2b42-2bfd-219a2f4aa21b
ms.date: 03/06/2019
localization_priority: Normal
---


# DoCmd.DoMenuItem method (Access)

Displays the appropriate menu or toolbar command for Microsoft Access.


## Syntax

_expression_.**DoMenuItem** (_MenuBar_, _MenuName_, _Command_, _Subcommand_, _Version_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MenuBar_|Required|**Variant**|Use the intrinsic constant **acFormBar** for the menu bar in Form view. For other views, use the number of the view in the _MenuBar_ argument list, as shown in the Macro window in previous versions of Microsoft Access (count down the list, starting from 0).|
| _MenuName_|Required|**Variant**|You can use one of the following intrinsic constants:<ul><li><p><b>acFile</b></p></li><li><p><b>acEditMenu</b></p></li><li><p><b>acRecordsMenu</b></p></li></ul><p>You can use <b>acRecordsMenu</b>  only for the Form view menu bar in Access version 2.0 and Access 95 databases. For other menus, use the number of the menu in the _MenuName_ argument list, as shown in the Macro window in previous versions of Access (count down the list, starting from 0).</p>|
| _Command_|Required|**Variant**|You can use one of the following intrinsic constants:<ul><li><p><b>acNew</b></p></li><li><p><b>acSaveForm</b></p></li><li><p><b>acSaveFormAs</b></p></li><li><p><b>acSaveRecord</b></p></li><li><p><b>acUndo</b></p></li><li><p><b>acCut</b></p></li><li><p><b>acCopy</b></p></li><li><p><b>acPaste</b></p></li><li><p><b>acDelete</b></p></li><li><p><b>acSelectRecord</b></p></li><li><p><b>acSelectAllRecords</b></p></li><li><p><b>acObjectRefresh</b></p></li></ul><p>For other commands, use the number of the command in the _Command_ argument list, as shown in the Macro window in previous versions of Access (count down the list, starting from 0).</p>|
| _Subcommand_|Optional|**Variant**|You can use one of the following intrinsic constants:<ul><li><p><b>acObjectVerb</b></p></li><li><p><b>acObjectUpdate</b></p></li></ul><p>The <b>acObjectVerb</b>  constant represents the first command on the submenu of the <b>Object</b> command on the <b>Edit</b> menu. The type of object determines the first command on the submenu. For example, this command is Edit for a Paintbrush object that can be edited.</p> <p>For other commands on submenus, use the number of the subcommand in the _Subcommand_ argument list, as shown in the Macro window in previous versions of Access (count down the list, starting from 0).</p>|
| _Version_|Optional|**Variant**|Use the intrinsic constant **acMenuVer70** for code written for Access 95 databases, the intrinsic constant **acMenuVer20** for code written for Access version 2.0 databases, and the intrinsic constant **acMenuVer1X** for code written for Access version 1.x databases. This argument is available only in Visual Basic.<br/><br/>**NOTE**: The default for this argument is **acMenuVer1X**, so that any code written for Access version 1.x databases will run unchanged. If you are writing code for a Access 95 or version 2.0 database and want to use the Access 95 or version 2.0 menu commands with the **DoMenuItem** method, you must set this argument to **acMenuVer70** or **acMenuVer20**.<br/><br/>Also, when you are counting down the lists for the _MenuBar_, _MenuName_, _Command_, and _Subcommand_ action arguments in the Macro window to get the numbers to use for the arguments in the **DoMenuItem** method, you must use the Access 95 lists if the _Version_ argument is **acMenuVer70**, the Access version 2.0 lists if the _Version_ argument is Version, and the Access version 1.x lists if _Version_ is **acMenuVer1X** (or blank).<br/><br/>**NOTE**: There is no **acMenuVer80** setting for this argument. You can't use the **DoMenuItem** method to display Access commands (although existing **DoMenuItem** methods in Visual Basic code will still work). Use the **[RunCommand](Access.Application.RunCommand.md)** method instead.|

## Remarks

> [!NOTE] 
> In Microsoft Access 97 and later, the **DoMenuItem** method was replaced by the **RunCommand** method. The **DoMenuItem** method is included in this version of Access only for compatibility with previous versions. When you run existing Visual Basic code containing a **DoMenuItem** method, Access will display the appropriate menu or toolbar command for Access 2000. However, unlike the DoMenuItem action in a macro, a **DoMenuItem** method in Visual Basic code isn't converted to a **RunCommand** method when you convert a database created in a previous version of Access.

Some commands from previous versions of Access aren't available in Access, and **DoMenuItem** methods that run these commands will cause an error when they're executed in Visual Basic. You must edit your Visual Basic code to replace or delete occurrences of such **DoMenuItem** methods.

The selections in the lists for the _MenuName_, _Command_, and _Subcommand_ action arguments in the Macro window depend on what you've selected for the previous arguments. You must use numbers or intrinsic constants that are appropriate for each _MenuBar_, _MenuName_, _Command_, and _Subcommand_ argument.

If you leave the _Subcommand_ argument blank but specify the _Version_ argument, you must include the _Subcommand_ argument's comma. If you leave the _Subcommand_ and _Version_ arguments blank, don't use a comma following the _Command_ argument.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
