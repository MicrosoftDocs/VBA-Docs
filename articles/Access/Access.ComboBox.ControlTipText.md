---
title: ComboBox.ControlTipText Property (Access)
keywords: vbaac10.chm11423
f1_keywords:
- vbaac10.chm11423
ms.prod: access
api_name:
- Access.ComboBox.ControlTipText
ms.assetid: 8562dde5-4bc7-92fb-347b-dd45e0eb413a
ms.date: 06/08/2017
---


# ComboBox.ControlTipText Property (Access)

You can use the  **ControlTipText** property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write **String**.


## Syntax

 _expression_. **ControlTipText**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

You set the  **ControlTipText** property by using a string expression up to 255 characters long.

For controls on forms, you can set the default for this property by using the default control style or the  **DefaultControl** property in Visual Basic.

You can set the  **ControlTipText** property in any view.

The  **ControlTipText** property provides an easy way to provide helpful information about controls on a form.

There are other ways to provide information about a form or a control on a form. You can use the  **StatusBarText** property to display information in the status bar about a control. To provide more extensive help for a form or control, use the **HelpFile** and **HelpContextID** properties.


## See also


#### Concepts


[ComboBox Object](Access.ComboBox.md)

