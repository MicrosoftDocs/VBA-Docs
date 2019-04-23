---
title: SmartDocument members (Office)
description: The SmartDocument property of the Document object in Word and the Workbook object in Excel returns a  SmartDocument object.
ms.prod: office
ms.assetid: 980de42d-6992-6107-a3fb-33e8c78da202
ms.date: 01/30/2019
localization_priority: Normal
---


# SmartDocument members (Office)

The **SmartDocument** property of the **Document** object in Microsoft Word and the **Workbook** object in Microsoft Excel returns a **SmartDocument** object.


## Methods

|Name|Description|
|:-----|:-----|
|[PickSolution](../../Office.SmartDocument.PickSolution.md)|Displays a dialog box that allows the user to choose an available XML expansion pack to attach to the active document in Microsoft Word or a workbook in Microsoft Excel.|
|[RefreshPane](../../Office.SmartDocument.RefreshPane.md)|Refreshes the **Document Actions** task pane for the active document in Microsoft Word or a workbook in Microsoft Excel.|


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.SmartDocument.Application.md)|Gets an **Application** object that represents the container application for the **SmartDocument** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Creator](../../Office.SmartDocument.Creator.md)|Gets a 32-bit integer that indicates the application in which the **SmartDocument** object was created. Read-only.|
|[SolutionID](../../Office.SmartDocument.SolutionID.md)|Gets or sets the ID, often a globally unique identifier (GUID), which identifies the XML expansion pack attached to the active document in Microsoft Word or workbook in Microsoft Excel. Read/write.|
|[SolutionURL](../../Office.SmartDocument.SolutionURL.md)|Gets or sets an absolute URL which provides the complete path to the XML expansion pack file attached to the active document in Microsoft Word or a workbook in Microsoft Excel. Read/write.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]