---
title: Form.Toolbar property (Access)
keywords: vbaac10.chm13386
f1_keywords:
- vbaac10.chm13386
ms.prod: access
api_name:
- Access.Form.Toolbar
ms.assetid: a004200c-5404-c3ba-f00d-591c0f0a545d
ms.date: 03/15/2019
localization_priority: Normal
---


# Form.Toolbar property (Access)

Specifies a custom toolbar to display for a form. Read/write **String**.


## Syntax

_expression_.**Toolbar**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

When opening a form in Microsoft Access that is part of a database that was created in an earlier version of Access, the specified toolbar will be displayed differently depending on the current settings of the **AllowFullMenus** and **AllowBuiltInToolbars** properties. 

If the **AllowFullMenus** and **AllowBuiltInToolbars** properties are set to **False**, the specified toolbar will replace the ribbon as the default set of commands available to the user. 

If the **AllowFullMenus** or **AllowBuiltInToolbars** property is set to **True**, the specified toolbar is displayed on the ribbon **Add-Ins** tab.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]