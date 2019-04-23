---
title: Application.ActivateMicrosoftApp method (Excel)
keywords: vbaxl10.chm133074
f1_keywords:
- vbaxl10.chm133074
ms.prod: excel
api_name:
- Excel.Application.ActivateMicrosoftApp
ms.assetid: e11d8165-5aad-2b1d-f9d1-797038d96afb
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.ActivateMicrosoftApp method (Excel)

Activates a Microsoft application. If the application is already running, this method activates the running application. If the application isn't running, this method starts a new instance of the application.


## Syntax

_expression_.**ActivateMicrosoftApp** (_Index_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **[XlMSApplication](Excel.XlMSApplication.md)**|Specifies the Microsoft application to activate.|

## Example

This example starts and activates Word.


```vb
Application.ActivateMicrosoftApp xlMicrosoftWord
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]