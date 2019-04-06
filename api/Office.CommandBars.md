---
title: CommandBars object (Office)
keywords: vbaof11.chm242000
f1_keywords:
- vbaof11.chm242000
ms.prod: office
api_name:
- Office.CommandBars
ms.assetid: 0e312e21-14ee-5055-d604-b66e61c53b47
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars object (Office)

A collection of **CommandBar** objects that represent the command bars in the container application.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Example

Use the **CommandBars** property to return the **CommandBars** collection. The following example displays in the Immediate window both the name and local name of each menu bar and toolbar, and it displays a value that indicates whether the menu bar or toolbar is visible.


```vb
For Each cbar in CommandBars 
    Debug.Print cbar.Name, cbar.NameLocal, cbar.Visible 
Next
```

<br/>

Use the **Add** method to add a new command bar to the collection. The following example creates a custom toolbar named **Custom1** and displays it as a floating toolbar.

```vb
Set cbar1 = CommandBars.Add(Name:="Custom1", Position:=msoBarFloating) 
cbar1.Visible = True
```

<br/>

Use enumName, where _index_ is the name or index number of a command bar, to return a single **CommandBar** object. The following example docks the toolbar named **Custom1** at the bottom of the application window.

```vb
CommandBars("Custom1").Position = msoBarBottom
```

> [!NOTE] 
> You can use the name or index number to specify a menu bar or toolbar in the list of available menu bars and toolbars in the container application. However, you must use the name to specify a menu, shortcut menu, or submenu (all of which are represented by **CommandBar** objects). If two or more custom menus or submenus have the same name, enumName returns the first one. To ensure that you return the correct menu or submenu, locate the popup control that displays that menu. Then apply the **CommandBar** property to the popup control to return the command bar that represents that menu.


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
