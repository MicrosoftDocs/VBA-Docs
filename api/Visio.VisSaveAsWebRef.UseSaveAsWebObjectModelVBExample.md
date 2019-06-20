---
title: "Using the Save as Web Page Object Model from Visual Basic: An example"
ms.prod: visio
ms.assetid: c5833ff8-45f3-ab67-3b16-09c60238965a
ms.date: 06/21/2019
localization_priority: Normal
---


# Using the Save as Web Page object model from Visual Basic: An example

To use the Save as Web Page API in your Visual Basic project, set a reference in your project to **Microsoft Visio 15.0 Save As Web Type Library**.

> [!NOTE] 
> In the Visual Basic Editor included with Visio, you can find the list of available references by choosing **References** on the **Tools** menu. In Visual Basic 6.0, you can find this list by choosing **References** on the **Project** menu.

The Save as Web Page model contains two classes: **[VisSaveAsWeb](Visio.VisSaveAsWeb.md)** and **[VisWebPageSettings](Visio.VisWebPageSettings.md)**, which implement the **IVisSaveAsWeb** and **IVisWebPageSettings** interfaces, respectively.

- A **VisSaveAsWeb** object implements the methods that perform the webpage creation process.  
- A **VisWebPageSettings** object contains the properties of your webpage project.
    
When you create a webpage and its supporting files (also called a webpage project), you'll typically follow these steps:

1. Use the **[SaveAsWebObject](visio.application.saveaswebobject.md)** property of the Visio **Application** object to get an instance of a **VisSaveAsWeb** object.
    
2. Use the **[WebPageSettings](visio.vissaveasweb.webpagesettings.md)** property of the **VisSaveAsWeb** object to get a reference to a **VisWebPageSettings** object, which you can use to get or set the webpage settings for your project.
    
3. Set the properties of the **VisWebPageSettings** object.
    
   > [!NOTE] 
   > You must always provide a target path for your files.

4. Call the **[AttachToVisioDoc](visio.vissaveasweb.attachtovisiodoc.md)** method to identify the document to save as a webpage. If you don't specify which document to save, the active drawing is saved.
    
5. Call the **[CreatePages](visio.vissaveasweb.createpages.md)** method to begin the Save as Web Page operation.
    

The following procedure shows how to open a new webpage project, set selected properties, and create the webpage files.

```vb
Public Sub SaveAsWeb () 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsoWebSettings As VisWebPageSettings 
 
 ' Get a VisSaveAsWeb object that 
 ' represents a new webpage project. 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 
 ' Get a VisWebPageSettings object. 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 ' Configure preferences. 
 With vsoWebSettings 
 .StartPage = 1 
 .EndPage = 2 
 .QuietMode = True 
 .TargetPath = "c:\your_folder_name\your_filename.htm" 
 End With 
 
 ' Create the pages. Because no particular document 
 ' is specified, the active drawing is saved. 
 vsoSaveAsWeb.CreatePages 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]