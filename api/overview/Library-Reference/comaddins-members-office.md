---
title: COMAddIns members (Office)
description: A collection of COMAddIn objects that provide information about a COM add-in registered in the Windows registry.
ms.prod: office
ms.assetid: 0fc908fa-0846-07ca-d2a2-4c87525ae719
ms.date: 01/30/2019
localization_priority: Normal
---


# COMAddIns members (Office)

A collection of **COMAddIn** objects that provide information about a COM add-in registered in the Windows registry.

## Methods

|Name|Description|
|:-----|:-----|
|[Item](../../Office.COMAddIns.Item.md)|Gets a member of the specified **COMAddIns** collection.|
|[Update](../../Office.COMAddIns.Update.md)|Updates the contents of the **COMAddIns** collection from the list of add-ins stored in the Windows registry.|

## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.COMAddIns.Application.md)|Gets an **Application** object that represents the container application for the **COMAddIns** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Count](../../Office.COMAddIns.Count.md)|Gets a count of the number of COM add-ins in the host application. Read-only.|
|[Creator](../../Office.COMAddIns.Creator.md)|Gets a 32-bit integer that indicates the application in which the **COMAddIns** object was created. Read-only.|
|[Parent](../../Office.COMAddIns.Parent.md)|Gets the **Parent** object for the **COMAddIns** object. Read-only.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]