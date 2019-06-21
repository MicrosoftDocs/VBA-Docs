---
title: VisSaveAsWeb.CreatePages method (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisSaveAsWeb.CreatePages
ms.assetid: 48094af2-55fb-9732-19bf-8a73827d1afb
ms.date: 06/21/2019
localization_priority: Normal
---


# VisSaveAsWeb.CreatePages method

Initiates webpage creation.


## Syntax

_expression_.**CreatePages**

_expression_ An expression that returns a **[VisSaveAsWeb](Visio.VisSaveAsWeb.md)** object.


## Return value

**Nothing**


## Remarks

Because the **VisSaveAsWeb** object uses the settings in its **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object to create the webpage, you should call the **CreatePages** method after you make any required changes to the **VisWebPageSettings** object.

To specify which document to save as a webpage, use the **[AttachToVisioDoc](Visio.VisSaveAsWeb.AttachToVisioDoc.md)** method. If no document is specified, Microsoft Visio saves the active document by default.


## Example

The following example shows how to open an existing file and save it as a webpage by using the Save as Web Page feature's default settings and the **AttachToVisioDoc** and **CreatePages** methods. Before running this example, replace `path\filename` with a valid path and file name for a Visio document to pass to the **Open** method. In addition, replace `targetpath\filename` with a valid target path and a file name for the webpage project files.


```vb
Public Sub CreatePages_Example () 
    Dim vsoSaveAsWeb As VisSaveAsWeb 
    Dim vsoWebSettings As VisWebPageSettings 
    Dim vsoDocument As Visio.Document
 
    Set vsoDocument = Application.Documents.Open("path\filename") 
    Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
    Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings
 
    vsoWebSettings.TargetPath = "targetpath\filename"
    
    With vsoSaveAsWeb
        .AttachToVisioDoc vsoDocument
        .CreatePages 
    End With
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]