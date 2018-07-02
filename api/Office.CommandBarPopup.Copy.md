---
title: CommandBarPopup.Copy Method (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.Copy
ms.assetid: d50fff50-00fd-e70f-d777-9bf1850cae37
ms.date: 06/08/2017
---


# CommandBarPopup.Copy Method (Office)

Copies a command bar popup control to an existing command bar.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. `Copy`( `_Bar_`, `_Before_` )

 _expression_ A variable that represents a [CommandBarPopup](./Office.CommandBarPopup.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Bar_|Optional|**Variant**|A  **CommandBar** object that represents the destination command bar. If this argument is omitted, the control is copied to the command bar where the control already exists.|
| _Before_|Optional|**Variant**|A number that indicates the position for the new control on the command bar. The new control will be inserted before the control at this position. If this argument is omitted, the control is copied to the end of the command bar.|

### Return Value

CommandBarControl


## See also


[CommandBarPopup Object](Office.CommandBarPopup.md)



[CommandBarPopup Object Members](./overview/commandbarpopup-members-office.md)

