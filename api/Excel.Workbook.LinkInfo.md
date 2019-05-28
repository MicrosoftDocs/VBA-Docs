---
title: Workbook.LinkInfo method (Excel)
keywords: vbaxl10.chm199108
f1_keywords:
- vbaxl10.chm199108
ms.prod: excel
api_name:
- Excel.Workbook.LinkInfo
ms.assetid: 016295a3-72c1-95b3-c259-8f286b12b73c
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.LinkInfo method (Excel)

Returns the link date and update status.


## Syntax

_expression_.**LinkInfo** (_Name_, _LinkInfo_, _Type_, _EditionRef_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the link.|
| _LinkInfo_|Required| **[XlLinkInfo](Excel.XlLinkInfo.md)**|The type of information to be returned.|
| _Type_|Optional| **Variant**|One of the constants of **[XlLinkInfoType](Excel.XlLinkInfoType.md)** specifying the type of link to return.|
| _EditionRef_|Optional| **Variant**|If the link is an edition, this argument specifies the edition reference as a string in R1C1 style. This argument is required if there's more than one publisher or subscriber with the same name in the workbook.|

## Return value

**Variant**


## Example

This example displays a message box if the link is updated automatically.

```vb
If ActiveWorkbook.LinkInfo( _ 
 "Word.Document|Document1!'!DDE_LINK1", xlUpdateState, _ 
 xlOLELinks) = 1 Then 
 MsgBox "Link updates automatically" 
End If
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]