---
title: VisSaveAsWeb Object (Visio Save As Web)
ms.prod: visio
ms.assetid: 48e19e11-9b41-42ec-84e9-c4aab7f08784
ms.date: 06/08/2017
localization_priority: Normal
---


# VisSaveAsWeb Object (Visio Save As Web)

Contains the webpage property settings and methods used when a Visio drawing is saved as a webpage. 


## Remarks

The **VisSaveAsWeb** object contains the methods and property settings that are used when a selected Visio drawing is saved as a webpage. The webpage project includes the following files:

- An HTML version of the drawing (including shape data, formerly called custom properties, and multiple drawing pages, if applicable).
    
- The supporting files associated with the project, for example, the graphics files (GIFs and JPGs), script files, data (XML) files, and cascading style sheet (CSS) files.
    
To set the properties for your webpage, use the **[WebPageSettings](Visio.WebPageSettings.md)** property of the **VisSaveAsWeb** object to get a **[VisWebPageSettings](visio.viswebpagesettings.object.visio.save.md)** object. After the properties are set, perform the following steps.


1. Call the **[AttachToVisioDoc](Visio.AttachToVisioDoc.md)** method to specify the drawing to be saved as a webpage. For example:
    
   ```vb
      vsoSaveAsWeb.AttachToVisioDoc _ 
    Application.Documents.Open("drive:\folder\drawingname.vdx")
   ```
   
   If you don't call this method, Visio creates the page from the active document by default.
    
2. Call the **[CreatePages](Visio.CreatePages.md)** method to create the webpage. For example:
    
   ```vb
      vsoSaveAsWeb.CreatePages vsoSaveAsWeb.CreatePages
   ```

You can control certain user interface behavior during page creation by using the **[SilentMode](Visio.SilentMode.md)** property or the **[QuietMode](Visio.QuietMode.md)** property of the **VisWebPageSettings** object.

The files created by the Save as Web Page feature are placed into the target path you specify, or a location you specify in the **[TargetPath](Visio.TargetPath.md)** property of the **VisWebPageSettings** object.

> [!NOTE] 
> You must specify a target path, or Visio will generate an error.

They can be organized as flat files or in a subfolder that has the same name as the drawing (see the **[StoreInFolder](Visio.StoreInFolder.md)** property of the **VisWebPageSettings** object).

> [!NOTE] 
> To view the **VisSaveAsWeb** class in the Object Browser, make sure that you have a reference to the Save As Web Page DLL in your project (in the Visual Basic Editor window, click **References**, on the **Tools** menu, and then select the **Microsoft Visio 15.0 SaveAsWeb Type Library** check box in the **Available References** list).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]