---
title: Application.ShortcutMenuBar property (Access)
keywords: vbaac10.chm12512
f1_keywords:
- vbaac10.chm12512
ms.prod: access
api_name:
- Access.Application.ShortcutMenuBar
ms.assetid: 6785320b-b50f-dcaa-3eae-13d378573613
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.ShortcutMenuBar property (Access)

You can use the **ShortcutMenuBar** property to specify the shortcut menu that appears when you right-click the specified object. Read/write **String**.


## Syntax

_expression_.**ShortcutMenuBar**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Remarks

When used with the **Application** object, the **ShortcutMenuBar** property enables you to display a custom shortcut menu as a global shortcut menu. However, if you've set the **ShortcutMenuBar** property for a form, form control, or report in the database, the custom shortcut menu of that object will be displayed in place of the database's global shortcut menu.

You can display a different custom shortcut menu for a specific form, form control, or report by setting its **ShortcutMenuBar** property to a different shortcut menu. When the form, form control, or report has the focus, the custom shortcut menu for that object is displayed when the user clicks the right mouse button; otherwise, the global shortcut menu for the database is displayed.

Shortcut menus aren't available to any object if the **AllowShortcutMenus** property is set to **False**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]