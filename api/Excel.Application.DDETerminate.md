---
title: Application.DDETerminate Method (Excel)
keywords: vbaxl10.chm183094
f1_keywords:
- vbaxl10.chm183094
ms.prod: excel
api_name:
- Excel.Application.DDETerminate
ms.assetid: f05adf6d-5714-12c4-39ce-af4bc31f4d32
ms.date: 06/08/2017
---


# Application.DDETerminate Method (Excel)

Closes a channel to another application.


## Syntax

 _expression_. `DDETerminate`( `_Channel_` )

 _expression_ A variable that represents an [Application](Excel.Application(Graph property).md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Channel_|Required| **Long**|The channel number returned by the  **[DDEInitiate](Excel.Application.DDEInitiate.md)** method.|

## Example

This example opens a channel to Word, opens the Word document Formletr.doc, and then sends the FilePrint command to WordBasic.


```vb
channelNumber = Application.DDEInitiate( _ 
 app:="WinWord", _ 
 topic:="C:\WINWORD\FORMLETR.DOC") 
Application.DDEExecute channelNumber, "[FILEPRINT]" 
Application.DDETerminate channelNumber
```


## See also


[Application Object](Excel.Application(object).md)

