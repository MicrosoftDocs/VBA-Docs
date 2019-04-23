---
title: Attachment.BorderColor property (Access)
keywords: vbaac10.chm13929
f1_keywords:
- vbaac10.chm13929
ms.prod: access
api_name:
- Access.Attachment.BorderColor
ms.assetid: cd43f030-f832-c58a-a374-67a349c3d499
ms.date: 02/07/2019
localization_priority: Normal
---


# Attachment.BorderColor property (Access)

You can use the **BorderColor** property to specify the color of a control's border. Read/write **Long**.


## Syntax

_expression_.**BorderColor**

_expression_ A variable that represents an **[Attachment](Access.Attachment.md)** object.


## Remarks

The **BorderColor** property setting is a numeric expression that corresponds to the color you want to use for a control's border.

You can set the default for this property by using a control's default control style or the **[DefaultControl](access.form.defaultcontrol.md)** property in Visual Basic.

A control's border color is visible only when its **SpecialEffect** property is set to Flat or Shadowed. If the **SpecialEffect** property is set to something other than Flat or Shadowed, setting the **BorderColor** property changes the **SpecialEffect** property setting to Flat.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]