---
title: Worksheet.ProtectScenarios property (Excel)
keywords: vbaxl10.chm174093
f1_keywords:
- vbaxl10.chm174093
ms.prod: excel
api_name:
- Excel.Worksheet.ProtectScenarios
ms.assetid: 7b0aacea-00f3-7f0a-2be1-693f0efbec88
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.ProtectScenarios property (Excel)

**True** if the worksheet scenarios are protected. Read-only **Boolean**.


## Syntax

_expression_.**ProtectScenarios**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example displays a message box if scenarios are protected on Sheet1.

```vb
If Worksheets("Sheet1").ProtectScenarios Then _ 
 MsgBox "Scenarios are protected on this worksheet."
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]