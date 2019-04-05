---
title: Application.IgnoreRemoteRequests property (Excel)
keywords: vbaxl10.chm133147
f1_keywords:
- vbaxl10.chm133147
ms.prod: excel
api_name:
- Excel.Application.IgnoreRemoteRequests
ms.assetid: 94515401-eb26-a2d8-5013-33f1f38b884f
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.IgnoreRemoteRequests property (Excel)

**True** if remote DDE requests are ignored. Read/write **Boolean**.


## Syntax

_expression_.**IgnoreRemoteRequests**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example sets the **IgnoreRemoteRequests** property to **True** so that remote DDE requests are ignored.


```vb
Application.IgnoreRemoteRequests = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]