---
title: BoundObjectFrame.ShortcutMenuBar property (Access)
keywords: vbaac10.chm10939
f1_keywords:
- vbaac10.chm10939
ms.prod: access
api_name:
- Access.BoundObjectFrame.ShortcutMenuBar
ms.assetid: 05f24e86-b02b-c55a-de10-0a6896ffefe0
ms.date: 02/08/2019
localization_priority: Normal
---


# BoundObjectFrame.ShortcutMenuBar property (Access)

You can use the **ShortcutMenuBar** property to specify the shortcut menu that appears when you right-click the specified object. Read/write **String**.


## Syntax

_expression_.**ShortcutMenuBar**

_expression_ A variable that represents a **[BoundObjectFrame](Access.BoundObjectFrame.md)** object.


## Remarks

The **ShortcutMenuBar** property applies only to controls on a form, and not to controls on a report.

You can also use the **ShortcutMenuBar** property to specify the menu bar macro that is used to display a shortcut menu for a datasheet, form, form control, or report. To display the built-in shortcut menu for a database, form, form control, or report by using a macro or Visual Basic, set the property to a zero-length string (" ").

When used with the **[Application](Access.Application.md)** object, the **ShortcutMenuBar** property enables you to display a custom shortcut menu as a global shortcut menu. However, if you've set the **ShortcutMenuBar** property for a form, form control, or report in the database, the custom shortcut menu of that object is displayed in place of the database's global shortcut menu. 

You can display a different custom shortcut menu for a specific form, form control, or report by setting its **ShortcutMenuBar** property to a different shortcut menu. When the form, form control, or report has the focus, the custom shortcut menu for that object is displayed when the user clicks the right mouse button; otherwise, the global shortcut menu for the database is displayed.

Shortcut menus aren't available to any object if the **AllowShortcutMenus** property is set to **False**.


## Example

The following example sets the **Suppliers_Toolbar** as the custom shortcut menu to display when the user clicks the right mouse button on the **Suppliers** form.


```vb
Forms("Suppliers").ShortcutMenuBar = "Suppliers_Toolbar"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]