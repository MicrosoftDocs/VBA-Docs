---
title: CommandBarPopup object (Office)
keywords: vbaof11.chm7000
f1_keywords:
- vbaof11.chm7000
ms.prod: office
api_name:
- Office.CommandBarPopup
ms.assetid: a8ae06a3-1d7b-a531-91df-756fafee5314
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarPopup object (Office)

Represents a popup control on a command bar.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Remarks

Every pop-up control contains a **CommandBar** object. To return the command bar from a pop-up control, apply the **CommandBar** property to the **CommandBarPopup** object.

Use **Controls**(_index_), where _index_ is the number of the control, to return a **CommandBarPopup** object. Note that the **[Type](office.msocontroltype.md)** property of the control must be **msoControlPopup**, **msoControlGraphicPopup**, **msoControlButtonPopup**, **msoControlSplitButtonPopup**, or **msoControlSplitButtonMRUPopup**.


## Example

You can also use the **FindControl** method to return a **CommandBarPopup** object. The following example searches all command bars for a **CommandBarPopup** object whose tag is **Graphics**.


```vb
Set myControl = Application.CommandBars.FindControl _ 
(Type:=msoControlPopup, Tag:="Graphics")
```


## See also

- [CommandBarPopup object members](overview/library-reference/commandbarpopup-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]