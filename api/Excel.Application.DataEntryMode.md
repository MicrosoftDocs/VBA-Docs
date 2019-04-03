---
title: Application.DataEntryMode property (Excel)
keywords: vbaxl10.chm133102
f1_keywords:
- vbaxl10.chm133102
ms.prod: excel
api_name:
- Excel.Application.DataEntryMode
ms.assetid: 1fd9f191-3aa5-2548-2d41-b9d2bc79654b
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.DataEntryMode property (Excel)

Returns or sets Data Entry mode, as shown in the following table. When in Data Entry mode, you can enter data only in the unlocked cells in the currently selected range. Read/write **Long**.


## Syntax

_expression_.**DataEntryMode**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

|Value|Description|
|:-----|:-----|
| **xlOn**|Data Entry mode is turned on.|
| **xlOff**|Data Entry mode is turned off.|
| **xlStrict**|Data Entry mode is turned on, and pressing Esc won't turn it off.|

## Example

This example turns off Data Entry mode if it's on.

```vb
If (Application.DataEntryMode = xlOn) Or _ 
 (Application.DataEntryMode = xlStrict) Then 
 Application.DataEntryMode = xlOff 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]