---
title: VisWebPageSettings object (Visio Save As Web)
ms.prod: visio
ms.assetid: 1f286540-2c46-4a2a-b133-2bfd6168db36
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings object (Visio Save As Web)

Contains the settings for the webpage.


## Remarks

The **VisWebPageSettings** object serves as a container for a webpage's properties.

Many of the properties of the **VisWebPageSettings** object correspond to the settings available in the **Save As** dialog box when a user chooses the **File** tab > **Export** > **Change File Type** > **Web Page (*.htm)** > **Save As**.

For example, the **PageTitle** property, which contains the title that appears in the title bar when a webpage is displayed in a browser, corresponds to the value in the **Page title** box in the **Set Page Title** dialog box (**Save As** dialog box > **Change Title**). 

Also, the **DispScreenRes** property corresponds to the value selected in the **Target Monitor** list on the **Advanced** tab of the **Save As Web Page** dialog box (**Save As** dialog box > **Save as type** list > **Web Page (\*.htm;\*.html)** > **Publish**).

When you want to create a webpage, use the **[WebPageSettings](Visio.VisSaveAsWeb.WebPageSettings.md)** property of the **VisSaveAsWeb** object to get a reference to the **VisWebPageSettings** object, which you can use to set the webpage's properties, as shown in the following example.

```vb
Public Sub VisWebPageSettingsObject_Example() 
 Dim vsoWebSettings As VisWebPageSettings 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 
 ' Query Visio for the VisSaveAsWeb object. 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 
 ' Get a WebPageSettings object. 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 
 ' Set the title that is displayed in the browser's title bar. 
 .PageTitle = "AccountingDeptOrgChart082501" 
 
 ' Prevent dialog boxes from appearing in the user interface. 
 .QuietMode = True 
 End With 
 
 ' If you do not call the AttachToVisioDoc method to 
 ' identify a specific document, Visio saves the 
 ' active document by default. 
 vsoSaveAsWeb.CreatePages 
End Sub
```

> [!NOTE] 
> To view the **VisWebPageSettings** class in the Object Browser, make sure that you have a reference to the Save As Web Page DLL in your project (in the Visual Basic Editor window, choose **References** on the **Tools** menu, and then select the **Microsoft Visio 15.0 Save As Web Type Library** check box in the **Available References** list).

## Methods

- [GetFormatName](Visio.VisWebPageSettings.GetFormatName.md)
- [GetPhysicalDimensions](Visio.VisWebPageSettings.GetPhysicalDimensions.md)
- [InitSettings](Visio.VisWebPageSettings.InitSettings.md)
- [ListFormats](Visio.VisWebPageSettings.ListFormats.md)
- [SaveSettings](Visio.VisWebPageSettings.SaveSettings.md)


## Properties

- [AltFormat](Visio.VisWebPageSettings.AltFormat.md)
- [DispScreenRes](Visio.VisWebPageSettings.DispScreenRes.md)
- [EndPage](Visio.VisWebPageSettings.EndPage.md)
- [FormatCount](Visio.VisWebPageSettings.FormatCount.md)
- [NavBar](Visio.VisWebPageSettings.NavBar.md)
- [OpenBrowser](Visio.VisWebPageSettings.OpenBrowser.md)
- [PageTitle](Visio.VisWebPageSettings.PageTitle.md)
- [PanAndZoom](Visio.VisWebPageSettings.PanAndZoom.md)
- [PriFormat](Visio.VisWebPageSettings.PriFormat.md)
- [PropControl](Visio.VisWebPageSettings.PropControl.md)
- [QuietMode](Visio.VisWebPageSettings.QuietMode.md)
- [Search](Visio.VisWebPageSettings.Search.md)
- [SecFormat](Visio.VisWebPageSettings.SecFormat.md)
- [SilentMode](Visio.VisWebPageSettings.SilentMode.md)
- [StartPage](Visio.VisWebPageSettings.StartPage.md)
- [StoreInFolder](Visio.VisWebPageSettings.StoreInFolder.md)
- [Stylesheet](Visio.VisWebPageSettings.Stylesheet.md)
- [TargetPath](Visio.VisWebPageSettings.TargetPath.md)
- [ThemeName](Visio.VisWebPageSettings.ThemeName.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]