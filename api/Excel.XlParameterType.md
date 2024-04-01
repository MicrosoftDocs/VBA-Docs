---
title: XlParameterType enumeration (Excel)
api_name:
- Excel.XlParameterType
ms.assetid: f6774f89-4992-2b7c-2dce-791fecafc1df
ms.date: 05/03/2019
ms.localizationpriority: medium
---


# XlParameterType enumeration (Excel)

Specifies how to determine the value of the parameter for the specified query table.

|Name|Value|Description|
|:-----|:-----|:-----|
| **xlConstant**|1|Uses the value specified by the _Value_ argument.|
| **xlPrompt**|0|Displays a dialog box that prompts the user for the value. The _Value_ argument specifies the text shown in the dialog box.|
| **xlRange**|2|Uses the value of the cell in the upper-left corner of the range. The _Value_ argument specifies a **[Range](Excel.Range(object).md)** object.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]