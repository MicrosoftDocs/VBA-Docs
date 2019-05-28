---
title: Workbook.ExclusiveAccess method (Excel)
keywords: vbaxl10.chm199099
f1_keywords:
- vbaxl10.chm199099
ms.prod: excel
api_name:
- Excel.Workbook.ExclusiveAccess
ms.assetid: 9b92ec4f-e256-7e01-6cd7-759a0d022813
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.ExclusiveAccess method (Excel)

Assigns the current user exclusive access to the workbook that's open as a shared list.


## Syntax

_expression_.**ExclusiveAccess**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Return value

**Boolean**


## Remarks

The **ExclusiveAccess** method saves any changes you've made to the workbook and requires other users who have the workbook open to save their changes to a different file.

If the specified workbook isn't open as a shared list, this method fails. To determine whether a workbook is open as a shared list, use the **[MultiUserEditing](Excel.Workbook.MultiUserEditing.md)** property.


## Example

This example determines whether the active workbook is open as a shared list. If it is, the example gives the current user exclusive access.

```vb
If ActiveWorkbook.MultiUserEditing Then 
 ActiveWorkbook.ExclusiveAccess 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]