---
title: Workbook.BreakLink method (Excel)
keywords: vbaxl10.chm199198
f1_keywords:
- vbaxl10.chm199198
ms.prod: excel
api_name:
- Excel.Workbook.BreakLink
ms.assetid: 1e9d70c1-908e-92eb-26b8-d6ac753cc9c2
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.BreakLink method (Excel)

Converts formulas linked to other Microsoft Excel sources or OLE sources to values.


## Syntax

_expression_.**BreakLink** (_Name_, _Type_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the link.|
| _Type_|Required| **[XlLinkType](Excel.XlLinkType.md)**|The type of link.|

## Example

In this example, Microsoft Excel converts the first link (an Excel link type) in the active workbook. This example assumes that at least one formula exists in the active workbook that links to another Excel source.

```vb
Sub UseBreakLink() 
 
 Dim astrLinks As Variant 
 
 ' Define variable as an Excel link type. 
 astrLinks = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks) 
 
 ' Break the first link in the active workbook. 
 ActiveWorkbook.BreakLink _ 
 Name:=astrLinks(1), _ 
 Type:=xlLinkTypeExcelLinks 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
