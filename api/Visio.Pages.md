---
title: Pages object (Visio)
keywords: vis_sdr.chm10195
f1_keywords:
- vis_sdr.chm10195
ms.prod: visio
api_name:
- Visio.Pages
ms.assetid: 45eec568-b5cc-5e80-ff5c-4dfa567efb5d
ms.date: 06/19/2019
localization_priority: Normal
---


# Pages object (Visio)

Includes a **[Page](Visio.Page.md)** object for each drawing page in a document.


## Remarks

To retrieve a **Pages** collection, use the **[Pages](visio.document.pages.md)** property of a **Document** object.

The default property of a **Pages** collection is **Item**.

The order of items in a **Pages** collection is significant: if there are _n_ foreground pages in a document, the first _n_ pages in its **Pages** collection are foreground pages and are in order. The remaining pages in the collection are the background pages of the document; these are in no particular order.

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this collection maps to the following types:

- **Microsoft.Office.Interop.Visio.IVPages**
    

## Events

- [AfterReplaceShapes](Visio.pages.afterreplaceshapes.md)
- [BeforePageDelete](Visio.Pages.BeforePageDelete.md)
- [BeforeReplaceShapes](Visio.pages.beforereplaceshapes.md)
- [BeforeSelectionDelete](Visio.Pages.BeforeSelectionDelete.md)
- [BeforeShapeDelete](Visio.Pages.BeforeShapeDelete.md)
- [BeforeShapeTextEdit](Visio.Pages.BeforeShapeTextEdit.md)
- [CalloutRelationshipAdded](Visio.Pages.CalloutRelationshipAdded.md)
- [CalloutRelationshipDeleted](Visio.Pages.CalloutRelationshipDeleted.md)
- [CellChanged](Visio.Pages.CellChanged.md)
- [ConnectionsAdded](Visio.Pages.ConnectionsAdded.md)
- [ConnectionsDeleted](Visio.Pages.ConnectionsDeleted.md)
- [ContainerRelationshipAdded](Visio.Pages.ContainerRelationshipAdded.md)
- [ContainerRelationshipDeleted](Visio.Pages.ContainerRelationshipDeleted.md)
- [ConvertToGroupCanceled](Visio.Pages.ConvertToGroupCanceled.md)
- [FormulaChanged](Visio.Pages.FormulaChanged.md)
- [GroupCanceled](Visio.Pages.GroupCanceled.md)
- [PageAdded](Visio.Pages.PageAdded.md)
- [PageChanged](Visio.Pages.PageChanged.md)
- [PageDeleteCanceled](Visio.Pages.PageDeleteCanceled.md)
- [QueryCancelConvertToGroup](Visio.Pages.QueryCancelConvertToGroup.md)
- [QueryCancelGroup](Visio.Pages.QueryCancelGroup.md)
- [QueryCancelPageDelete](Visio.Pages.QueryCancelPageDelete.md)
- [QueryCancelReplaceShapes](Visio.pages.querycancelreplaceshapes.md)
- [QueryCancelSelectionDelete](Visio.Pages.QueryCancelSelectionDelete.md)
- [QueryCancelUngroup](Visio.Pages.QueryCancelUngroup.md)
- [ReplaceShapesCanceled](Visio.pages.replaceshapescanceled.md)
- [SelectionAdded](Visio.Pages.SelectionAdded.md)
- [SelectionDeleteCanceled](Visio.Pages.SelectionDeleteCanceled.md)
- [ShapeAdded](Visio.Pages.ShapeAdded.md)
- [ShapeChanged](Visio.Pages.ShapeChanged.md)
- [ShapeDataGraphicChanged](Visio.Pages.ShapeDataGraphicChanged.md)
- [ShapeExitedTextEdit](Visio.Pages.ShapeExitedTextEdit.md)
- [ShapeLinkAdded](Visio.Pages.ShapeLinkAdded.md)
- [ShapeLinkDeleted](Visio.Pages.ShapeLinkDeleted.md)
- [ShapeParentChanged](Visio.Pages.ShapeParentChanged.md)
- [TextChanged](Visio.Pages.TextChanged.md)
- [UngroupCanceled](Visio.Pages.UngroupCanceled.md)

## Methods

- [Add](Visio.Pages.Add.md)
- [GetNames](Visio.Pages.GetNames.md)
- [GetNamesU](Visio.Pages.GetNamesU.md)

## Properties

- [Application](Visio.Pages.Application.md)
- [Count](Visio.Pages.Count.md)
- [Document](Visio.Pages.Document.md)
- [EventList](Visio.Pages.EventList.md)
- [Item](Visio.Pages.Item.md)
- [ItemFromID](Visio.Pages.ItemFromID.md)
- [ItemU](Visio.Pages.ItemU.md)
- [ObjectType](Visio.Pages.ObjectType.md)
- [PersistsEvents](Visio.Pages.PersistsEvents.md)
- [Stat](Visio.Pages.Stat.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
