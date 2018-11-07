---
title: Application.SheetFollowHyperlink Event (Excel)
keywords: vbaxl10.chm504093
f1_keywords:
- vbaxl10.chm504093
ms.prod: excel
api_name:
- Excel.Application.SheetFollowHyperlink
ms.assetid: 656e0ec6-64ea-1685-f088-a7e30bfaef38
ms.date: 06/08/2017
---


# Application.SheetFollowHyperlink Event (Excel)

Occurs when you click any hyperlink in Microsoft Excel. For worksheet-level events, see the Help topic for the  **[FollowHyperlink](Excel.Worksheet.FollowHyperlink.md)** event.


## Syntax

 _expression_. `SheetFollowHyperlink`( `_Sh_` , `_Target_` )

 _expression_ An expression that returns a [Application](Excel.Application(Graph property).md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The  **[Worksheet](Excel.Worksheet.md)** object that contains the hyperlink.|
| _Target_|Required| **Hyperlink**|The  **Hyperlink** object that represents the destination of the hyperlink.|

## See also


[Application Object](Excel.Application(object).md)

