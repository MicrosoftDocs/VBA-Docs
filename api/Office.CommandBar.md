---
title: CommandBar object (Office)
keywords: vbaof11.chm3000
f1_keywords:
- vbaof11.chm3000
ms.prod: office
api_name:
- Office.CommandBar
ms.assetid: 78603954-40aa-64cb-c407-2e0820d65231
ms.date: 01/03/2019
localization_priority: Priority
---


# CommandBar object (Office)

Represents a command bar in the container application. The **CommandBar** object is a member of the **CommandBars** collection.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Example

Use **CommandBars** (_index_), where _index_ is the name or index number of a command bar, to return a single **CommandBar** object. The following example steps through the collection of command bars to find the command bar named "Forms." If it finds this command bar, the example makes it visible and protects its docking state. In this example, the variable **cb** represents a **CommandBar** object.


```vb
foundFlag = False  
For Each cb In CommandBars 
    If cb.Name = "Forms" Then 
        cb.Protection = msoBarNoChangeDock 
        cb.Visible = True  
        foundFlag = True  
    End If 
Next cb 
If Not foundFlag Then 
    MsgBox "The collection does not contain a Forms command bar." 
End If
```

<br/>

You can use a name or index number to specify a menu bar or toolbar in the list of available menu bars and toolbars in the container application. However, you must use a name to specify a menu, shortcut menu, or submenu (all of which are represented by **CommandBar** objects). This example adds a new menu item to the bottom of the **Tools** menu. When chosen, the new menu item runs the procedure named "qtrReport."

```vb
Set newItem = CommandBars("Tools").Controls.Add(Type:=msoControlButton) 
With newItem 
    .BeginGroup = True  
    .Caption = "Make Report" 
    .FaceID = 0 
    .OnAction = "qtrReport" 
End With
```

<br/>

If two or more custom menus or submenus have the same name, **CommandBars(_index_)** returns the first one. To ensure that you return the correct menu or submenu, locate the pop-up control that displays that menu. Then apply the **CommandBar** property to the pop-up control to return the command bar that represents that menu. Assuming that the third control on the toolbar named **Custom Tools** is a pop-up control, this example adds the **Save** command to the bottom of that menu.

```vb
Set viewMenu = CommandBars("Custom Tools").Controls(3) 
viewMenu.Controls.Add ID:=3    'ID of Save command is 3
```


## See also

- [CommandBar object members](overview/library-reference/commandbar-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)
