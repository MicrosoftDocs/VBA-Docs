---
title: ModelConnection.CommandType property (Excel)
keywords: vbaxl10.chm922074
f1_keywords:
- vbaxl10.chm922074
ms.prod: excel
ms.assetid: 29343162-48b3-65c2-ccde-d780b81fd43d
ms.date: 05/01/2019
localization_priority: Normal
---


# ModelConnection.CommandType property (Excel)

Returns or sets one of the **[XlCmdType](excel.xlcmdtype.md)** enumeration constants. Read/write.


## Syntax

_expression_.**CommandType**

_expression_ A variable that represents a **[ModelConnection](Excel.modelconnection.md)** object.


## Remarks

For a **ModelConnection** object, this type will be set to either **xlCmdTable** or **xlCmdDAX**. The isolated connection **ThisWorkbookDataModel** to the data model will be of type **xlCmdCube**.


## Property value

**XLCMDTYPE**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]