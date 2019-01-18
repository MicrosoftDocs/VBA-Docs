---
title: CommandBars.AdaptiveMenus property (Office)
keywords: vbaof11.chm2013
f1_keywords:
- vbaof11.chm2013
ms.prod: office
api_name:
- Office.CommandBars.AdaptiveMenus
ms.assetid: 1b8c1a2a-9fe1-4148-6e03-5bf48f137d6f
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.AdaptiveMenus property (Office)

This property checks or unchecks the check box control for the option to show menus in Microsoft Office as full or personalized. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**AdaptiveMenus**

_expression_ A variable that represents a **[CommandBars](Office.CommandBars.md)** object.


## Remarks

Is **True** if adaptive menus are enabled for all applications in Microsoft Office. Read/write **Boolean**.

This control is set in any application by doing the following:

1.  On the **Tools** menu, select **Customize**.
    
2.  Select the **Options** tab.
    
3.  The **Always show full menus** option is located in the **Personalized Menus and Toolbars** section.
    

## Example

This example sets three options for all command bars in Office, including custom command bars and the controls on them.

```vb
With CommandBars 
    .LargeButtons = True  
    .DisplayFonts = True  
    .AdaptiveMenus = True  
End With
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]