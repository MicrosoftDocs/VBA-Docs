---
title: Workbook.SetLinkOnData method (Excel)
keywords: vbaxl10.chm199151
f1_keywords:
- vbaxl10.chm199151
ms.prod: excel
api_name:
- Excel.Workbook.SetLinkOnData
ms.assetid: b500a579-6e4c-5712-05cf-27c6393b3bcd
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.SetLinkOnData method (Excel)

Sets the name of a procedure that runs whenever a DDE link is updated.


## Syntax

_expression_.**SetLinkOnData** (_Name_, _Procedure_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the DDE/OLE link, as returned from the **[LinkSources](Excel.Workbook.LinkSources.md)** method.|
| _Procedure_|Optional| **Variant**|The name of the procedure to be run when the link is updated. This can be either a Microsoft Excel 4.0 macro or a Visual Basic procedure. Set this argument to an empty string ("") to indicate that no procedure should run when the link is updated.|

## Example

This example sets the name of the procedure that runs whenever the DDE link is updated.

```vb
ActiveWorkbook.SetLinkOnData _ 
 "WinWord|'C:\MSGFILE.DOC'!DDE_LINK1", _ 
 "my_Link_Update_Macro"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]