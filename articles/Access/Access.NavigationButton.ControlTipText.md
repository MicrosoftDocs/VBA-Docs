---
title: NavigationButton.ControlTipText Property (Access)
keywords: vbaac10.chm10480
f1_keywords:
- vbaac10.chm10480
ms.prod: access
api_name:
- Access.NavigationButton.ControlTipText
ms.assetid: 25ffbfdf-58e6-0bc0-a2b8-c4e94f85b864
ms.date: 06/08/2017
---


# NavigationButton.ControlTipText Property (Access)

You can use the  **ControlTipText** property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write **String**.


## Syntax

 _expression_. **ControlTipText**

 _expression_ A variable that represents a **NavigationButton** object.


## Remarks

You set the  **ControlTipText** property by using a string expression up to 255 characters long.

For controls on forms, you can set the default for this property by using the default control style or the  **DefaultControl** property in Visual Basic.

You can set the  **ControlTipText** property in any view.

The  **ControlTipText** property provides an easy way to provide helpful information about controls on a form.

There are other ways to provide information about a form or a control on a form. You can use the  **StatusBarText** property to display information in the status bar about a control. To provide more extensive help for a form or control, use the **HelpFile** and **HelpContextID** properties.


## See also


#### Concepts


[NavigationButton Object](Access.NavigationButton.md)

