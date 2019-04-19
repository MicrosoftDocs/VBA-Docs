---
title: Application.TransitionMenuKeyAction property (Excel)
keywords: vbaxl10.chm133219
f1_keywords:
- vbaxl10.chm133219
ms.prod: excel
api_name:
- Excel.Application.TransitionMenuKeyAction
ms.assetid: 8f278d3b-9902-597a-9e4d-7f2fc3f22469
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.TransitionMenuKeyAction property (Excel)

Returns or sets the action taken when the Microsoft Excel menu key is pressed. Can be either **xlExcelMenus** or **xlLotusHelp** (see the [Excel constants enumeration](excel.constants.md)). Read/write **Long**.


## Syntax

_expression_.**TransitionMenuKeyAction**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example sets the Microsoft Excel menu key to run Lotus 1-2-3 Help when it is pressed.

```vb
Application.TransitionMenuKeyAction = xlLotusHelp 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]