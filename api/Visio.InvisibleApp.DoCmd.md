---
title: InvisibleApp.DoCmd Method (Visio)
keywords: vis_sdr.chm17516190
f1_keywords:
- vis_sdr.chm17516190
ms.prod: visio
api_name:
- Visio.InvisibleApp.DoCmd
ms.assetid: d38786e8-5511-a489-937a-4fb9cd3fc502
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.DoCmd Method (Visio)

Performs the command that has the indicated command ID.


## Syntax

 _expression_. `DoCmd`( `_CommandID_` )

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _CommandID_|Required| **Integer**|The command to perform.|

## Return value

nothing


## Remarks

Constants for Microsoft Visio command IDs are declared by the Visio type library in  **[VisUICmds](Visio.visuicmds.md)** and are prefixed with **visCmd**.

The  **DoCmd** method works best with commands that display dialog boxes.

For a list of commands that can be used with the  **DoCmd** method, see the topic[DoCmd/DOCMD Commands ](../visio/Concepts/docmd-docmd-commands.md) in this Automation Reference.


## Example

The following macro shows how to use constants with the  **DoCmd** method. It opens a new document and displays the document stencil.


```vb
 
Public Sub DoCmd_Example() 
 
 Dim vsoDocument As Visio.Document 
 
 Set vsoDocument = Documents.Add("") 
 
 Visio.Application.DoCmd (visCmdWindowShowMasterObjects) 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]