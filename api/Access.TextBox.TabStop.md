---
title: TextBox.TabStop property (Access)
keywords: vbaac10.chm11064
f1_keywords:
- vbaac10.chm11064
ms.prod: access
api_name:
- Access.TextBox.TabStop
ms.assetid: ecb9ede6-e7a8-ca91-9ca3-3fad9de2ef90
ms.date: 02/26/2019
localization_priority: Normal
---


# TextBox.TabStop property (Access)

You can use the **TabStop** property to specify whether you can use the Tab key to move the focus to a control. Read/write **Boolean**.


## Syntax

_expression_.**TabStop**

_expression_ A variable that represents a **[TextBox](Access.TextBox.md)** object.


## Remarks

This property doesn't apply to toggle button controls when they appear in an option group. It applies only to the option group itself.

The **TabStop** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Yes|**True**|(Default) You can move the focus to the control by pressing the Tab key.|
|No|**False**|You can't move the focus to the control by pressing the Tab key.|

When you create a control on a form, Microsoft Access automatically assigns the control a position in the form's tab order. Each new control is placed last in the tab order. If you want to prevent a control from being available when you tab through the controls in a form, set the control's **TabStop** property to No.

In Form view, hidden or disabled controls remain in the tab order but are skipped when you move through the controls by pressing Tab, even if their **TabStop** properties are set to Yes.

As long as a control's **Enabled** property is set to Yes, you can click the control or use an access key to select it, regardless of its **TabStop** property setting. For example, you can set the **TabStop** property of a command button to No to prevent users from selecting the button by pressing Tab. However, they can still click the command button to choose it.


## Example

The following example disables the ability to move the focus to the **City** text box on the **Suppliers** form by using the Tab key.


```vb
Forms("Suppliers").Controls("City").TabStop = False
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]