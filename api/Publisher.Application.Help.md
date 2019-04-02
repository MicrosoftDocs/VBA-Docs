---
title: Application.Help method (Publisher)
keywords: vbapb10.chm131125
f1_keywords:
- vbapb10.chm131125
ms.prod: publisher
api_name:
- Publisher.Application.Help
ms.assetid: 37b51399-5897-4003-a0a9-9829a8adf8ed
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Help method (Publisher)

Displays online Help information.


## Syntax

 _expression_. **Help**(**_HelpType_**)

 _expression_ A variable that represents an  **Application** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|HelpType|Required| **PbHelpType**|The type of help to display.|

## Remarks

The HelpType parameter can be one of the following  **PbHelpType** constants declared in the Microsoft Publisher type library.



|Constant|Description|
|:-----|:-----|
| **pbHelp**|Displays the  **Help Topics** dialog box.|
| **pbHelpActiveWindow**|Displays Help describing the command associated with the active view or pane.|
| **pbHelpPSSHelp**| Displays product support information.|

Some of the constants listed above may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example displays a list of topics for troubleshooting printing problems.


```vb
Sub ShowPrintTroubleshooter() 
 Application.Help (HelpType:=pbHelpPrintTroubleshooter) 
End Sub
```


## See also


 [Application Object](Publisher.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]