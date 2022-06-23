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
> Interested in developing solutions that extend the Office experience across [multiple platforms](https://docs.microsoft.com/en-us/javascript/api/requirement-sets?view=common-js-preview)? Check out the new [Office Add-ins model](../../../../dev/add-ins/overview/office-add-ins.md).

### Graphical representation of Visio object model
 Please refer to the following links for more information on other notable Visio objects.    
**[Global](../../../api/visio.global.md "Global collection (Visio)")** &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; [ThisDocument](../../../visio/Concepts/about-the-thisdocument-object-visio.md "ThisDocumentobject (Visio)")    
├┉┉ [ActiveDocument](../../visio.application.activedocument.md "ActiveDocument object (Visio)")    
├┉┉ [ActivePage](../../../api/visio.global.activepage.md "ActivePage object (Visio)")    
├┉┉ [ActiveWindow](../../../api/visio.global.activewindow.md "ActiveWindow object (Visio)")    
├┉┉ [Application](../../../api/visio.application.md "Application object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;├ [ApplicationSettings](../../../api/visio.applicationsettings.md "ApplicationSettings object (Visio)")    
├┉┉├ [VBE](../../../api/visio.application.vbe.md "VBE object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;├ [CommandBars](../../../api/office.commandbars.md "CommandBars object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;├ [COMAddIns](../../../api/visio.application.comaddins.md "COMAddIns collection (Visio)")    
├┉┉├ **[Documents](../../../api/visio.global.documents.md "Documents collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│└ [Document](../../../api/visio.document.md "Document object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[Pages](../../../api/visio.pages.md "Pages collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │└ [Page](../../../api/visio.page.md "Page object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; ├ **[Shapes](../../../api/visio.shapes.md "Shapes collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│└ [Shape](../../../api/visio.shape.md "Shape object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; ├ **[Shapes](../../../api/visio.shapes.md "Sub-shapes collection (Visio)" )**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│ └ [Shape](../../../api/visio.shape.md "Sub-shape object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[Connects](../../../api/visio.connects.md "Connects collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│ └ [Connect](../../../api/visio.connect.md "Connect object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[Hyperlinks](../../../api/visio.hyperlinks.md "Hyperlinks collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│ └ [Hyperlink](../../../api/visio.hyperlink.md "Hyperlink object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ [Characters](../../../api/visio.characters.md "Characters object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; ├ [Section](../../../api/visio.section.md "Section object (Visio)") — [Row](../../../api/visio.row.md "Row object (Visio)") — [Cell](../../../api/visio.cell.md "Cell object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; └ **[Paths](../../../api/visio.paths.md "Paths collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; └ [Path](../../../api/visio.path.md "Path object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; └ [Curve](../../../api/visio.curve.md "Curve object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; ├ **[Connects](../../../api/visio.connects.md "Connects collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│ └ [ Connect](../../../api/visio.connect.md "Connect object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; ├ **[Layers](../../../api/visio.layers.md "Layers collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│ └ [ Layer](../../../api/visio.layer.md "Layer object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; └ **[OLEObjects](../../../api/visio.oleobjects.md "OLEObjects collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; └ [OLEObject](../../../api/visio.oleobject.md "OLEObject object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[Masters](../../../api/visio.masters.md "Masters collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │ └ [Master](../../../api/visio.master.md "Master object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; ├ **[Colors ](../../../api/visio.colors.md "Colors collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │ └ [Color](../../../api/visio.color.md "Color object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[Fonts ](../../../api/visio.fonts.md "Fonts collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │ └ [Font](../../../api/visio.font.md "Font object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[MasterShortcuts](../../../api/visio.mastershortcuts.md "MasterShortcuts collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │ └ [MasterShortcut](../../../api/visio.mastershortcut.md "MasterShortcut object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[OLEObjects](../../../api/visio.oleobjects.md "OLEObjects collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │ └ [OLEObject](../../../api/visio.oleobject.md "OLEObject object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├ **[Styles](../../../api/visio.styles.md "Styles collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp; │ └ [Style](../../../api/visio.style.md "Style object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;└ [VBProject](../../../api/visio.document.vbproject.md "VBProject object (Visio)")    
├┉┉├ **[Windows](../../../api/visio.global.windows.md "Windows collection (Visio)")**    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│└ [Window](../../../api/visio.window.md "Window object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;  ├ [Selection](../../../api/visio.selection.md "Selection object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;  ├ [Document](../../../api/visio.document.md "Document object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;  ├ [Master](../../../api/visio.master.md "Master object (Visio)")    
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│&nbsp;&nbsp;  └ [Page](../../../api/visio.page.md "Page object (Visio)")    
└┉┉├ **[Addons](../../../api/visio.global.addons.md "Addons collection (Visio)")**    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;│└ [Addon](../../../api/visio.document.md "Document object (Visio)")    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ├ **[EventList](../../../api/visio.eventlist.md "EventList collection (Visio)")**    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; │└ [Event](../../../api/visio.event.md "Event object (Visio)")    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; └ [UIObject](../../../api/visio.uiobject.md "UIObject object (Visio)")    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ├ **[AccelTables](../../../api/visio.acceltables.md "AccelTables collection (Visio)")**    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; │└ [AccelTable](../../../api/visio.acceltable.md "AccelTable object (Visio)")    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ├ **[MenuSets](../../../api/visio.menusets.md "MenuSets collection (Visio)")**    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; │└ [MenuSet](../../../api/visio.menuset.md "MenuSet object (Visio)")    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; └ **[ToolbarSets](../../../api/visio.toolbarsets.md "ToolbarSets collection (Visio)")**    
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; └ [ToolbarSet](../../../api/visio.toolbarset.md "ToolbarSet object (Visio)") 

## See also

- [Visio enumerations](../../visio(enumerations).md)
- [Getting started with VBA in Office](../../../Library-Reference/Concepts/getting-started-with-vba-in-office.md): Provides insight into how VBA programming can help to customize Office solutions.
- [What's new for VBA in Office 2019](../../../Library-Reference/Concepts/what-s-new-for-vba-in-office-2019.md): Lists the new VBA language elements for Office 2019.
- [What's new for VBA in Office 2016](../../../Library-Reference/Concepts/what-s-new-for-vba-in-office-2016.md): Lists the new VBA language elements for Office 2016.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
