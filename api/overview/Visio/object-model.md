---
title: Visio object model for Visual Basic for Applications (VBA)
description: This section of the Visio VBA Reference contains documentation for all the objects, properties, methods, and events contained in the Visio object model.
ms.prod: visio
ms.assetid: 166b707a-a5bf-42ae-7741-8ceb8a0ecfcc
ms.date: 01/22/2020
localization_priority: Normal
---


# Object model (Visio) 

This section of the Visio VBA Reference contains documentation for all the objects, properties, methods, and events contained in the Visio object model.

Use the table of contents in the left navigation to view the topics in this section.

> [!NOTE] 
> Interested in developing solutions that extend the Office experience across [multiple platforms](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability)? Check out the new [Office Add-ins model](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins).

### Graphical representation of Visio object model
 Please refer to the following links for more information on other notable Visio objects.    
**[Global](https://docs.microsoft.com/en-us/office/vba/api/visio.global "Global collection (Visio)")** &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; [ThisDocument](https://docs.microsoft.com/en-us/office/vba/visio/Concepts/about-the-thisdocument-object-visio "ThisDocumentobject (Visio)")    
├┉┉ [ActiveDocument](https://docs.microsoft.com/en-us/office/vba/api/visio.application.activedocument "ActiveDocument object (Visio)")    
├┉┉ [ActivePage](https://docs.microsoft.com/en-us/office/vba/api/visio.global.activepage "ActivePage object (Visio)")    
├┉┉ [ActiveWindow](https://docs.microsoft.com/en-us/office/vba/api/visio.global.activewindow "ActiveWindow object (Visio)")    
├┉┉ [Application](https://docs.microsoft.com/en-us/office/vba/api/visio.application "Application object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;├ [ApplicationSettings](https://docs.microsoft.com/en-us/office/vba/api/visio.applicationsettings "ApplicationSettings object (Visio)")    
├┉┉├ [VBE](https://docs.microsoft.com/en-us/office/vba/api/visio.application.vbe "VBE object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;├ [CommandBars](https://docs.microsoft.com/en-us/office/vba/api/office.commandbars "CommandBars object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;├ [COMAddIns](https://docs.microsoft.com/en-us/office/vba/api/visio.application.comaddins "COMAddIns collection (Visio)")    
├┉┉├ **[Documents](https://docs.microsoft.com/en-us/office/vba/api/visio.global.documents "Documents collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│└ [Document](https://docs.microsoft.com/en-us/office/vba/api/visio.document "Document object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[Pages](https://docs.microsoft.com/en-us/office/vba/api/visio.pages "Pages collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │└ [Page](https://docs.microsoft.com/en-us/office/vba/api/visio.page "Page object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; ├ **[Shapes](https://docs.microsoft.com/en-us/office/vba/api/visio.shapes "Shapes collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│└ [Shape](https://docs.microsoft.com/en-us/office/vba/api/visio.shape "Shape object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; ├ **[Shapes](https://docs.microsoft.com/en-us/office/vba/api/visio.shapes "Sub-shapes collection (Visio)" )**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│ └ [Shape](https://docs.microsoft.com/en-us/office/vba/api/visio.shape "Sub-shape object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[Connects](https://docs.microsoft.com/en-us/office/vba/api/visio.connects "Connects collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│ └ [Connect](https://docs.microsoft.com/en-us/office/vba/api/visio.connect "Connect object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[Hyperlinks](https://docs.microsoft.com/en-us/office/vba/api/visio.hyperlinks "Hyperlinks collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│ └ [Hyperlink](https://docs.microsoft.com/en-us/office/vba/api/visio.hyperlink "Hyperlink object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ [Characters](https://docs.microsoft.com/en-us/office/vba/api/visio.characters "Characters object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; ├ [Section](https://docs.microsoft.com/en-us/office/vba/api/visio.section "Section object (Visio)") — [Row](https://docs.microsoft.com/en-us/office/vba/api/visio.row "Row object (Visio)") — [Cell](https://docs.microsoft.com/en-us/office/vba/api/visio.cell "Cell object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; └ **[Paths](https://docs.microsoft.com/en-us/office/vba/api/visio.paths "Paths collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; └ [Path](https://docs.microsoft.com/en-us/office/vba/api/visio.path "Path object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; └ [Curve](https://docs.microsoft.com/en-us/office/vba/api/visio.curve "Curve object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; ├ **[Connects](https://docs.microsoft.com/en-us/office/vba/api/visio.connects "Connects collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│ └ [ Connect](https://docs.microsoft.com/en-us/office/vba/api/visio.connect "Connect object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; ├ **[Layers](https://docs.microsoft.com/en-us/office/vba/api/visio.layers "Layers collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│ └ [ Layer](https://docs.microsoft.com/en-us/office/vba/api/visio.layer "Layer object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; └ **[OLEObjects](https://docs.microsoft.com/en-us/office/vba/api/visio.oleobjects "OLEObjects collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; └ [OLEObject](https://docs.microsoft.com/en-us/office/vba/api/visio.oleobject "OLEObject object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[Masters](https://docs.microsoft.com/en-us/office/vba/api/visio.masters "Masters collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │ └ [Master](https://docs.microsoft.com/en-us/office/vba/api/visio.master "Master object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; ├ **[Colors ](https://docs.microsoft.com/en-us/office/vba/api/visio.colors "Colors collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │ └ [Color](https://docs.microsoft.com/en-us/office/vba/api/visio.color "Color object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[Fonts ](https://docs.microsoft.com/en-us/office/vba/api/visio.fonts "Fonts collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │ └ [Font](https://docs.microsoft.com/en-us/office/vba/api/visio.font "Font object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[MasterShortcuts](https://docs.microsoft.com/en-us/office/vba/api/visio.mastershortcuts "MasterShortcuts collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │ └ [MasterShortcut](https://docs.microsoft.com/en-us/office/vba/api/visio.mastershortcut "MasterShortcut object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[OLEObjects](https://docs.microsoft.com/en-us/office/vba/api/visio.oleobjects "OLEObjects collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │ └ [OLEObject](https://docs.microsoft.com/en-us/office/vba/api/visio.oleobject "OLEObject object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[Styles](https://docs.microsoft.com/en-us/office/vba/api/visio.styles "Styles collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │ └ [Style](https://docs.microsoft.com/en-us/office/vba/api/visio.style "Style object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;└ [VBProject](https://docs.microsoft.com/en-us/office/vba/api/visio.vbproject "VBProject object (Visio)")    
├┉┉├ **[Windows](https://docs.microsoft.com/en-us/office/vba/api/visio.global.windows "Windows collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│└ [Window](https://docs.microsoft.com/en-us/office/vba/api/visio.window "Window object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;  ├ [Selection](https://docs.microsoft.com/en-us/office/vba/api/visio.selection "Selection object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;  ├ [Document](https://docs.microsoft.com/en-us/office/vba/api/visio.document "Document object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;  ├ [Master](https://docs.microsoft.com/en-us/office/vba/api/visio.master "Master object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;  └ [Page](https://docs.microsoft.com/en-us/office/vba/api/visio.page "Page object (Visio)")    
└┉┉├ **[Addons](https://docs.microsoft.com/en-us/office/vba/api/visio.global.addons "Addons collection (Visio)")**    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│└ [Addon](https://docs.microsoft.com/en-us/office/vba/api/visio.document "Document object (Visio)")    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ├ **[EventList](https://docs.microsoft.com/en-us/office/vba/api/visio.eventlist "EventList collection (Visio)")**    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; │└ [Event](https://docs.microsoft.com/en-us/office/vba/api/visio.event "Event object (Visio)")    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; └ [UIObject](https://docs.microsoft.com/en-us/office/vba/api/visio.uiobject "UIObject object (Visio)")    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ├ **[AccelTables](https://docs.microsoft.com/en-us/office/vba/api/visio.acceltables "AccelTables collection (Visio)")**    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; │└ [AccelTable](https://docs.microsoft.com/en-us/office/vba/api/visio.acceltable "AccelTable object (Visio)")    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ├ **[MenuSets](https://docs.microsoft.com/en-us/office/vba/api/visio.menusets "MenuSets collection (Visio)")**    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; │└ [MenuSet](https://docs.microsoft.com/en-us/office/vba/api/visio.menuset "MenuSet object (Visio)")    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; └ **[ToolbarSets](https://docs.microsoft.com/en-us/office/vba/api/visio.toolbarsets "ToolbarSets collection (Visio)")**    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; └ [ToolbarSet](https://docs.microsoft.com/en-us/office/vba/api/visio.toolbarset "ToolbarSet object (Visio)")    

## See also

- [Visio enumerations](../../../api/visio(enumerations).md)
- [Getting started with VBA in Office](../../../Library-Reference/Concepts/getting-started-with-vba-in-office.md): Provides insight into how VBA programming can help to customize Office solutions.
- [What's new for VBA in Office 2019](../../../Library-Reference/Concepts/what-s-new-for-vba-in-office-2019.md): Lists the new VBA language elements for Office 2019.
- [What's new for VBA in Office 2016](../../../Library-Reference/Concepts/what-s-new-for-vba-in-office-2016.md): Lists the new VBA language elements for Office 2016.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
