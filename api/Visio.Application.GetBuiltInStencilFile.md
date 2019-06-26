---
title: Application.GetBuiltInStencilFile method (Visio)
keywords: vis_sdr.chm10062110
f1_keywords:
- vis_sdr.chm10062110
ms.prod: visio
api_name:
- Visio.Application.GetBuiltInStencilFile
ms.assetid: 2ae65aaa-d441-c7e8-3c8c-737bcca84738
ms.date: 06/26/2019
localization_priority: Normal
---


# Application.GetBuiltInStencilFile method (Visio)

Returns the file path to the specified built-in, hidden stencil used to populate certain galleries in the Microsoft Visio user interface.


## Syntax

_expression_.**GetBuiltInStencilFile** (_StencilType_, _MeasurementSystem_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _StencilType_|Required| **[VisBuiltInStencilTypes](Visio.VisBuiltInStencilTypes.md)**|The stencil to retrieve. Must be one of the **VisBuiltInStencilTypes** constants.|
| _MeasurementSystem_|Required| **[VisMeasurementSystem](Visio.vismeasurementsystem.md)**|The measurement system for the stencil.|

## Return value

**String**


## Example

The following Visual Basic for Applications (VBA) code sample shows how to use the **GetBuiltInStencilFile** method to open the built-in hidden container stencil, and to add one of the containers from that stencil to the active page to contain the selected shape or shapes. Before you run this code, be sure that there is a selected shape (or a selection of shapes) on the active page.

```vb
Public Sub GetBuiltInStencilFile_Example()

    Dim vsoDocument As Visio.Document
    Set vsoDocument = Application.Documents.OpenEx(Application.GetBuiltInStencilFile(visBuiltInStencilContainers, visMSUS), visOpenHidden)
    Application.ActivePage.DropContainer vsoDocument.Masters.ItemU("Container 1"), Application.ActiveWindow.Selection
    vsoDocument.Close

End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]