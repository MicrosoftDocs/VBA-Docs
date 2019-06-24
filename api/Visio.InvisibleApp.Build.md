---
title: InvisibleApp.Build property (Visio)
keywords: vis_sdr.chm17550515
f1_keywords:
- vis_sdr.chm17550515
ms.prod: visio
api_name:
- Visio.InvisibleApp.Build
ms.assetid: 912a1d47-e889-68b9-541b-12e9b9c36068
ms.date: 06/25/2019
localization_priority: Normal
---


# InvisibleApp.Build property (Visio)

Returns the build number of the running instance. Read-only.


## Syntax

_expression_.**Build**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


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