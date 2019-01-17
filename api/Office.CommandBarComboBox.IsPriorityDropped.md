---
title: CommandBarComboBox.IsPriorityDropped property (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.IsPriorityDropped
ms.assetid: c556f630-5e95-6d1a-4e94-0ecf5b20875a
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.IsPriorityDropped property (Office)

Gets **True** if the control is currently dropped from the menu or toolbar based on usage statistics and layout space. (Note that this is not the same as the control's visibility, as set by the **Visible** property). Read-only.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**IsPriorityDropped**

_expression_ A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


## Return value

Boolean


## Remarks

A control with **Visible** set to **True** will not be immediately visible on a personalized menu or toolbar if **IsPriorityDropped** is **True**.

To determine when to set **IsPriorityDropped** to **True** for a specific menu item, Microsoft Office maintains a total count of the number of times the menu item was used and a record of the number of different application sessions in which the user has used another menu item in the same menu as this menu item, without using the specific menu item. When this value reaches certain threshold values, the count is decremented. When the count reaches zero, **IsPriorityDropped** is set to **True**. Programmers cannot set the session value, the threshold value, or the **IsPriorityDropped** property. Programmers can, however, use the **AdaptiveMenus** property to disable adaptive menus for specific menus in an application.

To determine when to set **IsPriorityDropped** to **True** for a specific toolbar control, Office maintains a list of the order in which all the controls on that toolbar were last executed. A toolbar will always show as many controls as it has space to show, in the order of most recently used to least recently used. Controls with **Priority** set to 1 will always be shown and the toolbar will wrap rows, if necessary, to show these controls. Programmers can use the **Priority** property to ensure that specific toolbar controls are always shown, or to reposition toolbars so that they have enough space to display all of their controls.

You can use the following table to predict the number of sessions for which a menu item on a personalized menu will remain visible before the menu item's **IsPriorityDropped** property is set to **True**.

| Number of uses of the command bar control | Number of sessions of the application |
|:-----|:-----|
|0, 1|3|
|2|6|
|3|9|
|4, 5|12|
|6&ndash;8|17|
|9&ndash;13|23|
|14&ndash;24|29|
|25 or more|31|

## See also

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]