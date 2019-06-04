---
title: Application.ProductCode property (Publisher)
keywords: vbapb10.chm131105
f1_keywords:
- vbapb10.chm131105
ms.prod: publisher
api_name:
- Publisher.Application.ProductCode
ms.assetid: aacd5ff6-dad1-af86-f4e0-af9012ae93f8
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.ProductCode property (Publisher)

Returns a **String** indicating the Microsoft Publisher globally unique identifier (GUID). Read-only.


## Syntax

_expression_.**ProductCode**

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Return value

String


## Example

The following example displays the product code for Publisher.

```vb
MsgBox "The product code for Microsoft Publisher is " _ 
 & ProductCode
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]