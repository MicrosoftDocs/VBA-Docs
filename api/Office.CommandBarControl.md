---
title: CommandBarControl object (Office)
keywords: vbaof11.chm5000
f1_keywords:
- vbaof11.chm5000
ms.prod: office
api_name:
- Office.CommandBarControl
ms.assetid: b104ec00-beeb-a927-4b7b-108f4e3164f5
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarControl object (Office)

Represents a command bar control. The **CommandBarControl** object is a member of the **CommandBarControls** collection. The properties and methods of the **CommandBarControl** object are all shared by the **CommandBarButton**, **CommandBarComboBox**, and **CommandBarPopup** objects.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Remarks

When writing Visual Basic code to work with custom command bar controls, you use the **CommandBarButton**, **CommandBarComboBox**, and **CommandBarPopup** objects. 

When writing code to work with built-in controls in the container application that cannot be represented by one of those three objects, you use the **CommandBarControl** object. Use **Controls** (_index_), where _index_ is the index number of a control, to return a **CommandBarControl** object. (The **[Type](office.msocontroltype.md)** property of the control must be **msoControlLabel**, **msoControlExpandingGrid**, **msoControlSplitExpandingGrid**, **msoControlGrid**, or **msoControlGauge**). 

Variables declared as **CommandBarControl** can be assigned **CommandBarButton**, **CommandBarComboBox**, and **CommandBarPopup** values.


## Example

You can also use the **FindControl** method to return a **CommandBarControl** object. The following example searches for a control of type **msoControlGauge**; if it finds one, it displays the index number of the control and the name of the command bar that contains it. In this example, the variable _lbl_ represents a **CommandBarControl** object.


```vb
Set lbl = CommandBars.FindControl(Type:= msoControlGauge) 
If lbl Is Nothing Then 
    MsgBox "A control of type msoControlGauge was not found." 
Else 
    MsgBox "Control " & lbl.Index & " on command bar " _ 
        & lbl.Parent.Name & " is type msoControlGauge" 
End If
```


## See also

- [CommandBarControl object members](overview/library-reference/commandbarcontrol-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]