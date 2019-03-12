---
title: DoCmd.RunSavedImportExport method (Access)
keywords: vbaac10.chm5878
f1_keywords:
- vbaac10.chm5878
ms.prod: access
api_name:
- Access.DoCmd.RunSavedImportExport
ms.assetid: cb0ade9a-5cd4-1225-5231-8266fdfb3690
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.RunSavedImportExport method (Access)

Run a saved import or export specification.


## Syntax

_expression_.**RunSavedImportExport** (_SavedImportExportName_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SavedImportExportName_|Required|**Variant**| The name of a saved import or export specification to run.|

## Remarks

This method has the same effect as performing the following procedure in Access:

1. On the **External Data** tab, choose either **Saved Imports** or **Saved Exports**.
    
2. In the **Manage Data Tasks** dialog box, on the **Saved Imports** or **Saved Exports** tab (depending on your choice in the preceding step), choose the specification that you want to run.
    
3. Select **Run**. 
    
Before running the **RunSavedImportExport** method, make sure that the source and destination files exist, the source data is ready for importing, and the operation will not accidentally overwrite any data in your destination file.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
