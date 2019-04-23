---
title: CommandBarControls object (Office)
keywords: vbaof11.chm4000
f1_keywords:
- vbaof11.chm4000
ms.prod: office
api_name:
- Office.CommandBarControls
ms.assetid: 7ccae243-2870-95c2-1e08-140a3e638fe6
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarControls object (Office)

A collection of **[CommandBarControl](Office.CommandBarControl.md)** objects that represent the command bar controls on a command bar.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Example

Use the **Controls** property to return the **CommandBarControls** collection. The following example changes the caption of every control on the toolbar named **Standard** to the current value of the **Id** property for that control.

```vb
For Each ctl In CommandBars("Standard").Controls 
    ctl.Caption = CStr(ctl.Id) 
Next ctl
```

<br/>

Use the **Add** method to add a new command bar control to the **CommandBarControls** collection. This example adds a new, blank button to the command bar named **Custom**.

```vb
Set myBlankBtn = CommandBars("Custom").Controls.Add
```

<br/>

Use **Controls**(_index_), where _index_ is the caption or index number of a control, to return a **CommandBarControl**, **CommandBarButton**, **CommandBarComboBox**, or **CommandBarPopup** object. The following example copies the first control from the command bar named **Standard** to the command bar named **Custom**.

```vb
Set myCustomBar = CommandBars("Custom") 
Set myControl = CommandBars("Standard").Controls(1) 
myControl.Copy Bar:=myCustomBar, Before:=1
```


## See also

- [CommandBarControls object members](overview/library-reference/commandbarcontrols-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]