---
title: VisWebPageSettings object (Visio Save As Web)
ms.prod: visio
ms.assetid: 1f286540-2c46-4a2a-b133-2bfd6168db36
ms.date: 06/08/2017
localization_priority: Normal
---


# VisWebPageSettings object (Visio Save As Web)

Contains the settings for the webpage.


## Remarks

The  **VisWebPageSettings** object serves as a container for a webpage's properties.

Many of the properties of the  **VisWebPageSettings** object correspond to the settings available in the **Save As** dialog box when a user chooses the **File** tab, chooses **Export**, chooses  **Change File Type**, chooses  **Web Page (*.htm)**, and then chooses  **Save As**.

For example, the  **PageTitle** property, which contains the title that appears in the title bar when a webpage is displayed in a browser, corresponds to the value in the **Page title** box in the **Set Page Title** dialog box (in the **Save As** dialog box, choose **Change Title**). Also, the **DispScreenRes** property corresponds to the value selected in the **Target Monitor** list on the **Advanced** tab of the **Save As Web Page** dialog box (in the **Save As** dialog box, in the **Save as type** list, select **Web Page (\*.htm;\*.html)**, and then choose  **Publish**).

When you want to create a webpage, use the  **[WebPageSettings](Visio.VisSaveAsWeb.WebPageSettings.md)** property of the **VisSaveAsWeb** object to get a reference to the **VisWebPageSettings** object, which you can use to set the webpage's properties, as shown in the following example.




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


 **Note**  To view the  **VisWebPageSettings** class in the Object Browser, make sure that you have a reference to the Save As Web Page DLL in your project (in the Visual Basic Editor window, choose **References** on the **Tools** menu, and then select the **Microsoft Visio 15`.0 Save As Web Type Library** check box in the **Available References** list).

## Methods

- [GetFormatName](Visio.GetFormatName.md)
- [GetPhysicalDimensions](Visio.GetPhysicalDimensions.md)
- [InitSettings](Visio.InitSettings.md)
- [ListFormats](Visio.ListFormats.md)
- [SaveSettings](Visio.SaveSettings.md)

## Properties

- [AltFormat](Visio.AltFormat.md)
- [DispScreenRes](Visio.DispScreenRes.md)
- [EndPage](Visio.EndPage.md)
- [FormatCount](Visio.FormatCount.md)
- [NavBar](Visio.NavBar.md)
- [OpenBrowser](Visio.OpenBrowser.md)
- [PageTitle](Visio.PageTitle.md)
- [PanAndZoom](Visio.PanAndZoom.md)
- [PriFormat](Visio.PriFormat.md)
- [PropControl](Visio.PropControl.md)
- [QuietMode](Visio.QuietMode.md)
- [Search](Visio.Search.md)
- [SecFormat](Visio.SecFormat.md)
- [SilentMode](Visio.SilentMode.md)
- [StartPage](Visio.StartPage.md)
- [StoreInFolder](Visio.StoreInFolder.md)
- [Stylesheet](Visio.Stylesheet.md)
- [TargetPath](Visio.TargetPath.md)
- [ThemeName](Visio.ThemeName.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]