---
title: DataFeedConnection.SaveAsODC method (Excel)
keywords: vbaxl10.chm928088
f1_keywords:
- vbaxl10.chm928088
ms.prod: excel
ms.assetid: e66ff66c-9b19-a479-0afa-4f7e307113ac
ms.date: 03/28/2019
localization_priority: Normal
---


# DataFeedConnection.SaveAsODC method (Excel)

Saves the data feed connection as a Microsoft Office Data Connection file.


## Syntax

_expression_.**SaveAsODC** (_ODCFileName_, _Description_, _Keywords_)

_expression_ A variable that represents a **[DataFeedConnection](Excel.datafeedconnection.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ODCFileName_|Required|**String**|Location to save the file.|
| _Description_|Optional|**Variant**|Description that will be saved in the file.|
| _Keywords_|Optional|**Variant**|Space-separated keywords that can be used to search for this file.|

## Example

The following example saves the connection as an ODC file titled ODCFile. This example assumes data feed connection exists on the active worksheet. 

```vb
Sub UseSaveAsODC() 
 
   Application.ActiveWorkbook.Connections("Datafeed1").DataFeedConnection.SaveAsODC ("ODCFile")
 
End Sub
```


## Return value

**VOID**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]