---
title: Form.CloseButton property (Access)
keywords: vbaac10.chm13376
f1_keywords:
- vbaac10.chm13376
ms.prod: access
api_name:
- Access.Form.CloseButton
ms.assetid: c87e3752-0a77-3e5e-9c82-20effaf0af1e
ms.date: 03/12/2019
localization_priority: Normal
---


# Form.CloseButton property (Access)

Specifies whether the **Close** button on a form is enabled. Read/write **Boolean**.


## Syntax

_expression_.**CloseButton**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The **CloseButton** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Yes|**True**|(Default) The **Close** button is enabled.|
|No|**False**|The **Close** button is disabled, and the **Close** command isn't available on the **Control** menu.|

You can set the **CloseButton** property only in form Design view.

If you set the **CloseButton** property to No, the **Close** button remains visible but appears dimmed (grayed), and you must provide some other way to close the form; for example, a command button or custom menu command that runs a macro or event procedure that closes the form.

You can also close the form by pressing Alt+F4.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]