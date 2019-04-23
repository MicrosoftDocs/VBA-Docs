---
title: CommandBars.FindControl method (Office)
keywords: vbaof11.chm2007
f1_keywords:
- vbaof11.chm2007
ms.prod: office
api_name:
- Office.CommandBars.FindControl
ms.assetid: 07ec0c01-3cf4-3165-cfb2-c596b5e39abd
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.FindControl method (Office)

Gets a **CommandBarControl** object that fits a specified criteria.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**FindControl** (_Type_, _Id_, _Tag_, _Visible_)

_expression_ A variable that represents a **[CommandBars](Office.CommandBars.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**Variant**|The type of control.|
| _Id_|Optional|**Variant**|The identifier of the control.|
| _Tag_|Optional|**Variant**|The tag value of the control.|
| _Visible_|Optional|**Variant**|**True** to include only visible command bar controls in the search. The default value is **False**. Visible command bars include all visible toolbars and any menus that are open at the time the **FindControl** method is executed.|

## Return value

CommandBarControl


## Remarks

If the **CommandBars** collection contains two or more controls that fit the search criteria, **FindControl** returns the first control that's found. If no control that fits the criteria is found, **FindControl** returns **Nothing**.


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]