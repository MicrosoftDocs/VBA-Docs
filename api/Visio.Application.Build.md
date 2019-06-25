---
title: Application.Build property (Visio)
keywords: vis_sdr.chm10050515
f1_keywords:
- vis_sdr.chm10050515
ms.prod: visio
api_name:
- Visio.Application.Build
ms.assetid: 92fcdbe9-dfb1-cd20-4700-796bf7ca17f1
ms.date: 06/25/2019
localization_priority: Normal
---


# Application.Build property (Visio)

Returns the build number of the running instance. Read-only.


## Syntax

_expression_.**Build**

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Return value

Long


## Remarks

The format of the build number is described in the following table.

|Bits|Description|
|:-----|:-----|
|0 &ndash; 15|Internal build number|

The build number of the running instance is written to the **[BuildNumberCreated](visio.document.buildnumbercreated.md)** property when a new document is created, and to the **[BuildNumberEdited](visio.document.buildnumberedited.md)** property when a document is edited.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the **Build** property to get the build number of the running instance of Visio.

```vb
 
Public Sub Build_Example() 
 
 Debug.Print Application.Build 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]