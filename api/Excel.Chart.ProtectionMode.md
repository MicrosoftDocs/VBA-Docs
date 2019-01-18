---
title: Chart.ProtectionMode property (Excel)
keywords: vbaxl10.chm148092
f1_keywords:
- vbaxl10.chm148092
ms.prod: excel
api_name:
- Excel.Chart.ProtectionMode
ms.assetid: 5a9afe8c-df46-cbfe-d692-d4be8f2e505b
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ProtectionMode property (Excel)

 **True** if user-interface-only protection is turned on. To turn on user interface protection, use the **[Protect](Excel.Chart.Protect.md)** method with the _UserInterfaceOnly_ argument set to **True**. Read-only **Boolean**.


## Syntax

_expression_. `ProtectionMode`

_expression_ A variable that represents a [Chart](Excel.Chart-graph-object.md) object.


## Example

This example displays the status of the  **ProtectionMode** property.


```vb
MsgBox ActiveSheet.ProtectionMode
```


## See also


[Chart Object](Excel.Chart(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]