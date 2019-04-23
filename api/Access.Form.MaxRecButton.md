---
title: Form.MaxRecButton property (Access)
keywords: vbaac10.chm13488
f1_keywords:
- vbaac10.chm13488
ms.prod: access
api_name:
- Access.Form.MaxRecButton
ms.assetid: 6f5ea968-1f79-1fbc-86e1-fff034dcc827
ms.date: 03/13/2019
localization_priority: Normal
---


# Form.MaxRecButton property (Access)

You can use the **MaxRecButton** property to specify or determine if the maximum record limit button is available on the navigation bar of a form in Datasheet view or Form view. Read/write **Boolean**.


## Syntax

_expression_.**MaxRecButton**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

This property is only available for forms within a Microsoft Access project (.adp).

The default value is **True**.


## Example

This example makes the maximum record limit button on the **Order Entry** form unavailable to the user.

```vb
Forms("Order Entry").MaxRecButton = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]