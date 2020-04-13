---
title: Application.DDETerminate method (Word)
keywords: vbawd10.chm158335290
f1_keywords:
- vbawd10.chm158335290
ms.prod: word
api_name:
- Word.Application.DDETerminate
ms.assetid: c469656c-edf8-3ce2-b09b-0883faba8943
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DDETerminate method (Word)

Closes the specified dynamic data exchange (DDE) channel to another application.


## Syntax

_expression_. `DDETerminate`( `_Channel_` )

_expression_ A variable that represents an **[Application](Word.Application.md)** object.  Optional.


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


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]