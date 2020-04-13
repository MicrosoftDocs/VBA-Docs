---
title: Global.DDETerminate method (Word)
keywords: vbawd10.chm163119418
f1_keywords:
- vbawd10.chm163119418
ms.prod: word
api_name:
- Word.Global.DDETerminate
ms.assetid: 2502d0a7-c90b-1169-7b7b-a5d2b26445a6
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.DDETerminate method (Word)

Closes the specified dynamic data exchange (DDE) channel to another application.


## Syntax

_expression_. `DDETerminate`( `_Channel_` )

_expression_ A variable that represents a '[Global](Word.Global.md)' object. Optional.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Channel_|Required| **Long**|The channel number returned by the **DDEInitiate** method.|


## Example

This example creates a new workbook in Microsoft Excel and then terminates the DDE conversation.


```vb
Dim lngChannel As Long 
 
lngChannel = DDEInitiate(App:="Excel", Topic:="System") 
DDEExecute Channel:=lngChannel, Command:="[New(1)]" 
DDETerminate Channel:=lngChannel
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]