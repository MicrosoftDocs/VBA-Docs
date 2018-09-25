---
title: InvisibleApp.ShapeLinkAdded Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.ShapeLinkAdded
ms.assetid: b7637b06-8d74-04db-c4ab-a64d9b92f8a6
ms.date: 06/08/2017
---


# InvisibleApp.ShapeLinkAdded Event (Visio)

Occurs after a shape is linked to a data row.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

Private Sub  _expression_ _'ShapeLinkAdded'(**_ByVal Shape As [IVSHAPE]_** , **_ByVal DataRecordsetID As Long_** , **_ByVal DataRowID As Long_**)

 _expression_ An expression that returns a [InvisibleApp](./Visio.InvisibleApp.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape that is linked to data.|
| _DataRecordsetID_|Required| **Long**|The ID of the data recordset containing the data row linked to the shape.|
| _DataRowID_|Required| **Long**|The ID of the data row linked to the shape.|

## Remarks

The  **ShapeLinkAdded** event is one of a group of events for which the **EventInfo** property of the **Application** object contains extra information.

When the  **ShapeLinkAdded** event is fired, the **EventInfo** property returns the following string:

 `/DataRecordsetID = n /DataRowID = m`

where  _n_ and _m_ represent the IDs of the data recordset and data row, respectively, associated with the event.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


