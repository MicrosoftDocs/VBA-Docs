---
title: CommandBar.NameLocal property (Office)
keywords: vbaof11.chm3011
f1_keywords:
- vbaof11.chm3011
ms.prod: office
api_name:
- Office.CommandBar.NameLocal
ms.assetid: 3afad045-aaf8-8775-574e-faaccde7d270
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.NameLocal property (Office)

Gets the name of a built-in command bar as it's displayed in the language version of the container application, or returns or sets the name of a custom command bar. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**NameLocal**

_expression_ A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Remarks

> [!NOTE]
> If you attempt to set this property for a built-in command bar, an error occurs.

The local name of a built-in command bar is displayed in the title bar (when the command bar isn't docked) and in the list of available command bars, wherever that list is displayed in the container application.

If you change the value of the **LocalName** property for a custom command bar, the value of **Name** changes as well, and vice versa.


## Example

This example displays the name and localized name of the first command bar in the container application.


```vb
With CommandBars(1) 
    MsgBox "The name of the command bar is " & .Name 
    MsgBox "The localized name of the command bar is " & .NameLocal 
End With
```


## See also

- [CommandBar object members](overview/library-reference/commandbar-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]