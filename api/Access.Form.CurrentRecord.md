---
title: Form.CurrentRecord property (Access)
keywords: vbaac10.chm13473
f1_keywords:
- vbaac10.chm13473
ms.prod: access
api_name:
- Access.Form.CurrentRecord
ms.assetid: a682d187-0b9a-2fc3-3443-f2dcd6df4ca2
ms.date: 03/12/2019
localization_priority: Normal
---


# Form.CurrentRecord property (Access)

You can use the **CurrentRecord** property to identify the current record in the recordset being viewed on a form. Read/write **Long**.


## Syntax

_expression_.**CurrentRecord**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Microsoft Access sets this property to a **Long Integer** value that represents the current record number displayed on a form.

The **CurrentRecord** property is read-only in Form view and Datasheet view. It's not available in Design view.

The value specified by this property corresponds to the value shown in the record number box found in the lower-left corner of the form.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
