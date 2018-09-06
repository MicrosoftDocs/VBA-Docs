---
title: IRTDUpdateEvent.HeartbeatInterval Property (Excel)
keywords: vbaxl10.chm500002
f1_keywords:
- vbaxl10.chm500002
ms.prod: excel
api_name:
- Excel.IRTDUpdateEvent.HeartbeatInterval
ms.assetid: 45a3df85-59c1-fedb-e94b-8f011601fc72
ms.date: 06/08/2017
---


# IRTDUpdateEvent.HeartbeatInterval Property (Excel)

Returns or sets a  **Long** for the interval between updates for real-time data. Read/write.


## Syntax

 _expression_. `HeartbeatInterval`

 _expression_ A variable that represents an [IRTDUpdateEvent](Excel.IRTDUpdateEvent.md) object.


## Remarks

Setting the  **HeartbeatInterval** property to -1 will result in the **[Heartbeat](Excel.IRtdServer.Heartbeat.md)** method not being called.


 **Note**  The heartbeat interval cannot be set below 15,000 milliseconds, due to the standard 15-second time out.


## See also


[IRTDUpdateEvent Object](Excel.IRTDUpdateEvent.md)

