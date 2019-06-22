---
title: Window.SelectedDataRecordset property (Visio)
keywords: vis_sdr.chm11660245
f1_keywords:
- vis_sdr.chm11660245
ms.prod: visio
api_name:
- Visio.Window.SelectedDataRecordset
ms.assetid: 89c6b4ba-fb39-8932-1fe0-9a8aa2cbaef0
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.SelectedDataRecordset property (Visio)

Gets or sets the data recordset that is displayed on the active tab of the  **External Data Window** in the Microsoft Visio user interface (UI). Read/write.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_. `SelectedDataRecordset`

 _expression_ An expression that returns a **[Window](Visio.Window.md)** object.


## Return value

DataRecordset


## Remarks

The  **SelectedDataRecordset** property works only when the **Window** object represents the **External Data Window**. Calling the property on any other window type results in an error. The  **External Data Window** must already be displayed in the Visio UI before you call **SelectedDataRecordset**.

When you set the  **SelectedDataRecordset** property, the **DataRecordset** object you pass must not have been added with the **visDataRecordsetNoExternalDataUI** flag set.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]