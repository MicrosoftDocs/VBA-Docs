---
title: Workbook.Connections property (Excel)
description: Learn how use Workbook.Connections property to return a Connections object (Excel)
keywords: vbaxl10.chm199235
f1_keywords:
- vbaxl10.chm199235
api_name:
- Excel.Workbook.Connections
ms.assetid: 9c4f4ba7-dd4b-0bc2-65b7-16455014097f
ms.date: 07/20/2021
ms.localizationpriority: medium
---


# Workbook.Connections property (Excel)

Returns a [Connections](Excel.Connections.md) object that is a container for connections between the workbook and data sources such as ODBC, OLEDB, etc., that can refresh the data without prompting the user. Read-only.


## Syntax

_expression_.**Connections**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

The following example refreshes the OBDC and OLEDB connections of the active workbook.

```vb
ActiveWorkbook.Connections(1).ODBCConnection.Refresh 
ActiveWorkbook.Connections(1).OLEDBConnection.Refresh 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
