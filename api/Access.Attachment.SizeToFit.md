---
title: Attachment.SizeToFit method (Access)
keywords: vbaac10.chm13907
f1_keywords:
- vbaac10.chm13907
ms.prod: access
api_name:
- Access.Attachment.SizeToFit
ms.assetid: 9e9b8a65-79ba-9fda-08d8-9b5444678228
ms.date: 02/07/2019
localization_priority: Normal
---


# Attachment.SizeToFit method (Access)

You can use the **SizeToFit** method to size a control so that it fits the text or image that it contains.


## Syntax

_expression_.**SizeToFit**

_expression_ A variable that represents an **[Attachment](Access.Attachment.md)** object.


## Remarks

For example, you can apply the **SizeToFit** method to a command button that is too small to display all the text in its **Caption** property.

You can apply the **SizeToFit** method to controls only in form Design view or report Design view.

The **SizeToFit** method makes a control larger or smaller, depending on the size of the text or image that it contains.

You can use the **SizeToFit** method in conjunction with the **[CreateControl](Access.Application.CreateControl.md)** method to size new controls that you have created in code.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]