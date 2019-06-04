---
title: Application.Help method (Publisher)
keywords: vbapb10.chm131125
f1_keywords:
- vbapb10.chm131125
ms.prod: publisher
api_name:
- Publisher.Application.Help
ms.assetid: 37b51399-5897-4003-a0a9-9829a8adf8ed
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.Help method (Publisher)

Displays online Help information.


## Syntax

_expression_.**Help** (_HelpType_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_HelpType_|Required| **[PbHelpType](publisher.pbhelptype.md)**| The type of help to display. Can be one of the **PbHelpType** constants.|

## Remarks

Some of the **PbHelpType** constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example displays a list of topics for troubleshooting printing problems.

```vb
Sub ShowPrintTroubleshooter() 
 Application.Help (HelpType:=pbHelpPrintTroubleshooter) 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]