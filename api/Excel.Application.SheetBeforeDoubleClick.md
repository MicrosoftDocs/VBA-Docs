---
title: Application.SheetBeforeDoubleClick Event (Excel)
keywords: vbaxl10.chm504075
f1_keywords:
- vbaxl10.chm504075
ms.prod: excel
api_name:
- Excel.Application.SheetBeforeDoubleClick
ms.assetid: 969394a3-2c87-36a5-2d64-521bad8849be
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SheetBeforeDoubleClick Event (Excel)

Occurs when any worksheet is double-clicked, before the default double-click action.


## Syntax

_expression_. `SheetBeforeDoubleClick`( `_Sh_` , `_Target_` , `_Cancel_` )

 _expression_ An expression that returns a [Application](Excel.Application-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**| A **[Worksheet](Excel.Worksheet.md)** object that represents the sheet.|
| _Target_|Required| **Range**|The cell nearest to the mouse pointer when the double-click occurred.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the default double-click action isn't performed when the procedure is finished.|

## Remarks

This event doesn't occur on chart sheets.


## See also


[Application Object](Excel.Application(object).md)

