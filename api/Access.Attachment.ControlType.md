---
title: Attachment.ControlType property (Access)
keywords: vbaac10.chm13914
f1_keywords:
- vbaac10.chm13914
ms.prod: access
api_name:
- Access.Attachment.ControlType
ms.assetid: f660ca13-59f0-efae-8e6b-7449662a15c2
ms.date: 02/07/2019
localization_priority: Normal
---


# Attachment.ControlType property (Access)

You can use the **ControlType** property in Visual Basic to determine the type of control on a form or report. Read/write **Byte**.


## Syntax

_expression_.**ControlType**

_expression_ A variable that represents an **[Attachment](Access.Attachment.md)** object.


## Remarks

The **ControlType** property setting is an **[AcControlType](Access.AcControlType.md)** constant that specifies the control type.

The **ControlType** property can only be set by using Visual Basic in form Design view or report Design view, but it can be read in all views.

The **ControlType** property is also used to specify the type of control to create when you are using the **[CreateControl](Access.Application.CreateControl.md)** method.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]