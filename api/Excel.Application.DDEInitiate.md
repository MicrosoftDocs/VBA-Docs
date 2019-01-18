---
title: Application.DDEInitiate method (Excel)
keywords: vbaxl10.chm132090
f1_keywords:
- vbaxl10.chm132090
ms.prod: excel
api_name:
- Excel.Application.DDEInitiate
ms.assetid: 4b14e2ee-d7b0-a028-42a7-0809fa381f7e
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DDEInitiate method (Excel)

Opens a DDE channel to an application.


## Syntax

_expression_. `DDEInitiate`( `_App_` , `_Topic_` )

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _App_|Required| **String**|The application name.|
| _Topic_|Required| **String**|Describes something in the application to which you are opening a channel ? usually a document of that application.|

## Return value

Long


## Remarks

If successful, the  **DDEInitiate** method returns the number of the open channel. All subsequent DDE functions use this number to specify the channel.


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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]