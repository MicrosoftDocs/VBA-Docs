---
title: ModelConnection.CommandType property (Excel)
keywords: vbaxl10.chm922074
f1_keywords:
- vbaxl10.chm922074
ms.prod: excel
ms.assetid: 29343162-48b3-65c2-ccde-d780b81fd43d
ms.date: 06/08/2017
---


# ModelConnection.CommandType property (Excel)

Returns or sets one of the [xlCmdType enumeration (Excel)](Excel.xlCmdType.md) constants. Read/Write


## Syntax

 _expression_. `CommandType`

 _expression_ A variable that represents a [ModelConnection object (Excel)](Excel.modelconnection.md) object.


## Remarks

For a  **ModelConnection** object, this type will be set to either **xlCmdTable** or **xlCmdDAX**. The isolated connection **ThisWorkbookDataModel** to the Data Model will be of type **xlCmdCube**.


## Property value

 **XLCMDTYPE**


## See also



[ModelConnection Object](Excel.modelconnection.md)

