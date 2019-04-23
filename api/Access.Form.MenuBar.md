---
title: Form.MenuBar property (Access)
keywords: vbaac10.chm13385
f1_keywords:
- vbaac10.chm13385
ms.prod: access
api_name:
- Access.Form.MenuBar
ms.assetid: b9e6b6f6-5e60-271d-67c4-6697cb294671
ms.date: 03/13/2019
localization_priority: Normal
---


# Form.MenuBar property (Access)

Specifies a custom menu to display for a form. Read/write **String**.


## Syntax

_expression_.**MenuBar**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

When opening a form in Microsoft Access that is part of a database that was created in an earlier version of Access, the specified menu bar will be displayed differently depending on the current settings of the **AllowFullMenus** and **AllowBuiltInToolbars** properties. 

If the **AllowFullMenus** and **AllowBuiltInToolbars** properties are set to **False**, the specified menu bar will replace the ribbon as the default set of commands available to the user. 

If the **AllowFullMenus** or **AllowBuiltInToolbars** property is set to **True**, the specified menu bar is displayed on the ribbon **Add-Ins** tab.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]