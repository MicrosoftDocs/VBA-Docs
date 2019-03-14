---
title: Report.DisplayOnSharePointSite property (Access)
keywords: vbaac10.chm13873
f1_keywords:
- vbaac10.chm13873
ms.prod: access
api_name:
- Access.Report.DisplayOnSharePointSite
ms.assetid: 4e13b1e9-3b79-d073-fb51-848fdc2dcada
ms.date: 03/15/2019
localization_priority: Normal
---


# Report.DisplayOnSharePointSite property (Access)

Gets or sets whether the specified report can be made available as a view on a Microsoft SharePoint Foundation site. Read/write **Byte**.


## Syntax

_expression_.**DisplayOnSharePointSite**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

The **DisplayOnSharePointSite** property uses the following settings.

|Setting|Description|
|:-----|:-----|
|Do Not Display|Do not make the report an available view. |
|Follow Table Setting|(Default) Make the report an available view if the report's parent table is configured to be displayed as a view.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]