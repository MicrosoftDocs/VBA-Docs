---
title: Workbook.AutoUpdateSaveChanges property (Excel)
keywords: vbaxl10.chm199079
f1_keywords:
- vbaxl10.chm199079
ms.prod: excel
api_name:
- Excel.Workbook.AutoUpdateSaveChanges
ms.assetid: 06f9951d-a17a-bf88-4f6e-65835eb112f8
ms.date: 05/25/2019
localization_priority: Normal
---


# Workbook.AutoUpdateSaveChanges property (Excel)

**True** if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. **False** if changes aren't posted (this workbook is still synchronized with changes made by other users). The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**AutoUpdateSaveChanges**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

The **[AutoUpdateFrequency](Excel.Workbook.AutoUpdateFrequency.md)** property must be set to a value from 5 to 1440 for this property to take effect.


## Example

This example causes changes to the shared workbook to be posted to other users whenever the workbook is automatically updated.

```vb
ActiveWorkbook.AutoUpdateSaveChanges = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]