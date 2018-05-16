---
title: Attachment.SetFocus Method (Access)
keywords: vbaac10.chm13909
f1_keywords:
- vbaac10.chm13909
ms.prod: access
api_name:
- Access.Attachment.SetFocus
ms.assetid: 53e9c9f3-f1b8-f68d-8e9a-8b15ab4a3e83
ms.date: 06/08/2017
---


# Attachment.SetFocus Method (Access)

The  **SetFocus** method moves the focus to the specified form, the specified control on the active form, or the specified field on the active datasheet.


## Syntax

 _expression_. **SetFocus**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

You can use the  **SetFocus** method when you want a particular field or control to have the focus so that all user input is directed to this object.

Other properties can be set only when a control does not have the focus. For example, you can't set a control's  **Visible** or **Enabled** properties to **False** (0) when that control has the focus.

You can also use the  **SetFocus** method to navigate in a form according to certain conditions. For example, if the user selects **Not applicable** for the first of a set of questions on a questionnaire form, your Visual Basic code might then automatically skip the questions in that set and move the focus to the first control in the next set of questions.

You can move the focus only to a visible control or form. A form and controls on a form aren't visible until the form's  **Load** event has finished. Therefore, if you use the **SetFocus** method in a form's **Load** event to move the focus to that form, you must use the **Repaint** method before the **SetFocus** method.

You can't move the focus to a control if its  **Enabled** property is set to **False**. You must set a control's **Enabled** property to **True** (-1) before you can move the focus to that control. You can, however, move the focus to a control if its **Locked** property is set to **True**.

If a form contains controls for which the  **Enabled** property is set to **True**, you can't move the focus to the form itself. You can only move the focus to controls on the form. In this case, if you try to use **SetFocus** to move the focus to a form, the focus is set to the control on the form that last received the focus.


## See also


#### Concepts


[Attachment Object](Access.Attachment.md)

