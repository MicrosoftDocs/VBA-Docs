---
title: Workbook.AutoSaveOn property (Excel)
keywords: vbaxl10.chm199287
f1_keywords:
- vbaxl10.chm199287
api_name:
- Excel.Workbook.AutoSaveOn
ms.date: 05/25/2019
ms.localizationpriority: medium
---


# Workbook.AutoSaveOn property (Excel)

**True** if the edits in the workbook are automatically saved. Read/write **Boolean**.

## Syntax

_expression_.**AutoSaveOn**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.

## Remarks

When a new workbook is created, the default value for the **AutoSaveOn** property is **False**, the property is disabled, and the user's changes will need to be saved manually. However, if the workbook is hosted in the cloud (that is, OneDrive, OneDrive for Business, or SharePoint Online), the **AutoSaveOn** property defaults to **True** and the edits in the specified workbook are automatically saved. If a cloud-hosted workbook is shared with other users, their changes will also be automatically merged into the user's local copy when **AutoSaveOn** is **True**.

The following table shows examples of **AutoSaveOn** behavior.

|AutoSaveOn toggle state|Set AutoSaveOn to True|Set AutoSaveOn to False|
|:-----|:-----|:-----|
|`AutoSaveOn == True`|No-op|`AutoSaveOn` turned off|
|`AutoSaveOn == False`|`AutoSaveOn` turned on|No-op|
|Disabled|Error|Error|

For more information about AutoSave, see [How AutoSave impacts add-ins and macros](../Library-Reference/Concepts/how-autosave-impacts-addins-and-macros.md).

## Example

This example notifies you whether the workbook is set to be automatically saved.

```vb
Sub UseAutoSaveOn()
    MsgBox "This workbook is being saved automatically: " & ActiveWorkbook.AutoSaveOn
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]