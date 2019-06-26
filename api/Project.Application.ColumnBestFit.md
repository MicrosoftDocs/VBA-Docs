---
title: Application.ColumnBestFit method (Project)
keywords: vbapj.chm2037
f1_keywords:
- vbapj.chm2037
ms.prod: project-server
api_name:
- Project.Application.ColumnBestFit
ms.assetid: 51f96761-33ab-d2e3-7a1e-c8266bdaa7a1
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ColumnBestFit method (Project)

Sets the width of a column to the width of its widest item.


## Syntax

_expression_. `ColumnBestFit`( `_Column_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Column_|Optional|**Long**|A number that specifies the column to adjust. Columns are numbered from left to right, starting with 1. If  **Column** is omitted, Project adjusts the width of the column that contains the active cell.|

## Return value

 **Boolean**


## Example

The following example adjusts the widths of the first five columns in the active table.


```vb
Sub BestFitFirstFiveCols() 
 
    Dim I As Integer ' Index used in For...Next loop. 
 
    For I = 1 To 5 
          ColumnBestFit Column:=I 
    Next I 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]