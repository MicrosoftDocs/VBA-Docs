---
title: DataRecordset.BeforeDataRecordsetDelete Event (Visio)
keywords: vis_sdr.chm16462030
f1_keywords:
- vis_sdr.chm16462030
ms.prod: visio
api_name:
- Visio.DataRecordset.BeforeDataRecordsetDelete
ms.assetid: 6cb35848-51fe-653d-6cb3-a91e324bc6f3
ms.date: 06/08/2017
localization_priority: Normal
---


# DataRecordset.BeforeDataRecordsetDelete Event (Visio)

Occurs before a  **DataRecordset** object is deleted from the **DataRecordsets** collection.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

Private Sub  _expression_ _'BeforeDataRecordsetDelete'(**_ByVal DataRecordset As IVDATARECORDSET_**)

 _expression_ An expression that returns a [DataRecordset](./Visio.DataRecordset.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DataRecordset_|Required| **[IVDATARECORDSET]**|The data recordset that is going to be deleted.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]