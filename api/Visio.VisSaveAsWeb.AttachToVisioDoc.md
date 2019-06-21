---
title: VisSaveAsWeb.AttachToVisioDoc method (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisSaveAsWeb.AttachToVisioDoc
ms.assetid: ed2aba12-21b0-d953-8f5b-0634255f03b5
ms.date: 06/21/2019
localization_priority: Normal
---


# VisSaveAsWeb.AttachToVisioDoc method

Indicates which document to save as a webpage.


## Syntax

_expression_.**AttachToVisioDoc** (_docObj_)

_expression_ An expression that returns a **[VisSaveAsWeb](Visio.VisSaveAsWeb.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_docObj_|Required| **Document**|An Automation object that supports the **IVDocument** interface.|

## Return value

**Nothing**


## Remarks

Microsoft Visual Basic programs can pass a Visio **Document** object to this method because objects created from the Visio **Document** class support the **IVDocument** interface.

The **AttachToVisioDoc** method queries the **IUnknown** interface for the presence of the **IVDocument** interface.


## Example

The following example shows how to open an existing file and save it as a webpage by using the Save as Web Page feature's default settings and the **AttachToVisioDoc** and **[CreatePages](Visio.VisSaveAsWeb.CreatePages.md)** methods. Before running this example, replace `path\filename` with a valid path and file name for a Visio document to pass to the **Open** method. In addition, replace `targetpath\filename` with a valid target path and a file name for the webpage project files.

```vb

Public Sub AttachToVisioDoc_Example () 
    Dim vsoSaveAsWeb As VisSaveAsWeb 
    Dim vsoWebSettings As VisWebPageSettings 
    Dim vsoDocument As Visio.Document
 
    Set vsoDocument = Application.Documents.Open("<variable>path\filename</variable>") 
    Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject
    Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings
 
    vsoWebSettings.TargetPath = "<variable>targetpath\filename</variable>"

    With vsoSaveAsWeb
        .AttachToVisioDoc vsoDocument
        .CreatePages 
    End With
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]