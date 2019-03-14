---
title: Form.ServerFilterByForm property (Access)
keywords: vbaac10.chm13483
f1_keywords:
- vbaac10.chm13483
ms.prod: access
api_name:
- Access.Form.ServerFilterByForm
ms.assetid: f9f8f28e-b67e-1f4e-a70b-c66169fca250
ms.date: 03/15/2019
localization_priority: Normal
---


# Form.ServerFilterByForm property (Access)

You can use the **ServerFilterByForm** property to specify or determine whether a form is opened in the Server Filter By Form window. Read/write **Boolean**.


## Syntax

_expression_.**ServerFilterByForm**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The default value is **False**.

You can remove a filter by using Visual Basic to set the **ServerFilterByForm** property to **False**.

> [!NOTE] 
> The **ServerFilterByForm** property setting is ignored if the form's record source is a stored procedure.


## Example

The following example enables the **Order Lookup** form to be opened in a Microsoft Access data project in the Server Filter By Form window.

```vb
Forms("Order Lookup").ServerFilterByForm = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]