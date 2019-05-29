---
title: Workbook.RunAutoMacros method (Excel)
keywords: vbaxl10.chm199143
f1_keywords:
- vbaxl10.chm199143
ms.prod: excel
api_name:
- Excel.Workbook.RunAutoMacros
ms.assetid: 85dfdadf-75e6-437d-fb7a-e17681a69b35
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.RunAutoMacros method (Excel)

Runs the Auto_Open, Auto_Close, Auto_Activate, or Auto_Deactivate macro attached to the workbook. This method is included for backward compatibility. For new Visual Basic code, you should use the **Open**, **Activate**, and **Deactivate** events and the **Close** method instead of these macros.


## Syntax

_expression_.**RunAutoMacros** (_Which_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Which_|Required| **[XlRunAutoMacro](Excel.XlRunAutoMacro.md)**|Specifies the automatic macro to run.|


## Example

This example opens the workbook Analysis.xls and then runs its Auto_Open macro.

```vb
Workbooks.Open "ANALYSIS.XLS" 
ActiveWorkbook.RunAutoMacros xlAutoOpen
```

<br/>

This example runs the Auto_Close macro for the active workbook and then closes the workbook.

```vb
With ActiveWorkbook 
 .RunAutoMacros xlAutoClose 
 .Close 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
