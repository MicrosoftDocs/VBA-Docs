---
title: OLEDBConnection.RefreshOnFileOpen Property (Excel)
keywords: vbaxl10.chm794086
f1_keywords:
- vbaxl10.chm794086
ms.prod: excel
api_name:
- Excel.OLEDBConnection.RefreshOnFileOpen
ms.assetid: 09a0b59d-7a6e-65a6-d72a-14460d787ed9
ms.date: 06/08/2017
---


# OLEDBConnection.RefreshOnFileOpen Property (Excel)

 **True** if the connection is automatically updated each time the workbook is opened. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_. `RefreshOnFileOpen`

 _expression_ A variable that represents an [OLEDBConnection](Excel.OLEDBConnection.md) object.


## Remarks

The connections are not automatically refreshed when you open the workbook by using the  **[Open](Excel.Workbooks.Open.md)** method in Visual Basic. Use the **[Refresh](Excel.OLEDBConnection.Refresh.md)** method to refresh the data after the workbook is open.


## See also


[OLEDBConnection Object](Excel.OLEDBConnection.md)

