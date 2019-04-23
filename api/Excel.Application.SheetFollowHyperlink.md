---
title: Application.SheetFollowHyperlink event (Excel)
keywords: vbaxl10.chm504093
f1_keywords:
- vbaxl10.chm504093
ms.prod: excel
api_name:
- Excel.Application.SheetFollowHyperlink
ms.assetid: 656e0ec6-64ea-1685-f088-a7e30bfaef38
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.SheetFollowHyperlink event (Excel)

Occurs when you click any hyperlink in Microsoft Excel. For worksheet-level events, see the Help topic for the **[FollowHyperlink](Excel.Worksheet.FollowHyperlink.md)** event.


## Syntax

_expression_.**SheetFollowHyperlink** (_Sh_, _Target_)

_expression_ An expression that returns an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The **[Worksheet](Excel.Worksheet.md)** object that contains the hyperlink.|
| _Target_|Required| **Hyperlink**|The **[Hyperlink](excel.hyperlink.md)** object that represents the destination of the hyperlink.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]