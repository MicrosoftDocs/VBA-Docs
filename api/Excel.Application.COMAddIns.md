---
title: Application.COMAddIns property (Excel)
keywords: vbaxl10.chm133246
f1_keywords:
- vbaxl10.chm133246
ms.prod: excel
api_name:
- Excel.Application.COMAddIns
ms.assetid: d51f3373-ba5d-20b4-7557-246a6fcf89c3
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.COMAddIns property (Excel)

Returns the **[COMAddIns](Office.COMAddIns.md)** collection for Microsoft Excel, which represents the currently installed COM add-ins. Read-only.


## Syntax

_expression_.**COMAddIns**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the number of COM add-ins that are currently installed.

```vb
Set objAI = Application.COMAddIns 
MsgBox "Number of COM add-ins available:" & _ 
    objAI.Count
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]