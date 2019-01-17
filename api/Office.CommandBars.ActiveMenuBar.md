---
title: CommandBars.ActiveMenuBar property (Office)
keywords: vbaof11.chm2002
f1_keywords:
- vbaof11.chm2002
ms.prod: office
api_name:
- Office.CommandBars.ActiveMenuBar
ms.assetid: 8f341f53-418c-6d05-ac0b-e45a6b2baa0d
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.ActiveMenuBar property (Office)

Gets a **CommandBar** object that represents the active menu bar in the container application. Read-only.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**ActiveMenuBar**

_expression_ A variable that represents a **[CommandBars](Office.CommandBars.md)** object.


## Example

This example adds a temporary pop-up control named **Custom** to the end of the active menu bar, and adds a control named **Import** to the popup control.


```vb
Set myMenuBar = CommandBars.ActiveMenuBar 
Set newMenu = myMenuBar.Controls.Add(Type:=msoControlPopup, Temporary:=True) 
newMenu.Caption = "Custom" 
Set ctrl1 = newMenu.CommandBar.Controls _ 
    .Add(Type:=msoControlButton, Id:=1) 
With ctrl1 
    .Caption = "Import" 
    .TooltipText = "Import" 
    .Style = msoButtonCaption 
End With
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]