---
title: VisSaveAsWeb.AttachToVisioDoc Method (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.AttachToVisioDoc
ms.assetid: ed2aba12-21b0-d953-8f5b-0634255f03b5
ms.date: 06/08/2017
localization_priority: Normal
---


# VisSaveAsWeb.AttachToVisioDoc Method (Visio Save As Web)

Indicates which document to save as a webpage.


## Syntax

_expression_.**AttachToVisioDoc**(**_docObj_**)

 _expression_ An expression that returns a  **[VisSaveAsWeb](overview/Visio.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|docObj |Required| **Document**|An Automation object that supports the  **IVDocument** interface.|

## Return value

 **Nothing**


## Remarks

Microsoft Visual Basic programs can pass a Visio  **Document** object to this method because objects created from the Visio **Document** class support the **IVDocument** interface.

The  **AttachToVisioDoc** method queries the **IUnknown** interface for the presence of the **IVDocument** interface.


## Example

The following example shows how to open an existing file and save it as a webpage by using the Save as Web Page feature's default settings and the  **AttachToVisioDoc** and **[CreatePages](Visio.CreatePages.md)** methods. Before running this example, replace _path\filename_ with a valid path and file name for a Visio document to pass to the **Open** method. In addition, replace _targetpath\filename_ with a valid target path and a file name for the Web page project files.


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