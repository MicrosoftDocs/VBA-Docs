---
title: TextBox.EnterKeyBehavior property (Access)
keywords: vbaac10.chm11054,vbaac10.chm4343
f1_keywords:
- vbaac10.chm11054,vbaac10.chm4343
ms.prod: access
api_name:
- Access.TextBox.EnterKeyBehavior
ms.assetid: b7830316-a1aa-ddc1-094f-5976c5298bc1
ms.date: 03/26/2019
localization_priority: Normal
---


# TextBox.EnterKeyBehavior property (Access)

You can use the **EnterKeyBehavior** property to specify what happens when you press Enter in a text box control in Form view or Datasheet view. Read/write **Boolean**.

## Syntax

_expression_.**EnterKeyBehavior**

_expression_ A variable that represents a **[TextBox](Access.TextBox.md)** object.


## Remarks

For example, you can use this property if you have a control bound to a **Memo** field in a table to make entering multiple-line text easier. If you don't set this property to New Line In Field, you must press Ctrl+Enter to enter a new line in the text box.

You can also set the default for this property by setting a control's **[DefaultControl](access.form.defaultcontrol.md)** property in Visual Basic.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]