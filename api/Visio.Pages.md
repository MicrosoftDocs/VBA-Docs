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

- [fterReplaceShapes](Visio.pages.afterreplaceshapes.md)
- [eforePageDelete](Visio.Pages.BeforePageDelete.md)
- [eforeReplaceShapes](Visio.pages.beforereplaceshapes.md)
- [eforeSelectionDelete](Visio.Pages.BeforeSelectionDelete.md)
- [eforeShapeDelete](Visio.Pages.BeforeShapeDelete.md)
- [eforeShapeTextEdit](Visio.Pages.BeforeShapeTextEdit.md)
- [alloutRelationshipAdded](Visio.Pages.CalloutRelationshipAdded.md)
- [alloutRelationshipDeleted](Visio.Pages.CalloutRelationshipDeleted.md)
- [ellChanged](Visio.Pages.CellChanged.md)
- [onnectionsAdded](Visio.Pages.ConnectionsAdded.md)
- [onnectionsDeleted](Visio.Pages.ConnectionsDeleted.md)
- [ontainerRelationshipAdded](Visio.Pages.ContainerRelationshipAdded.md)
- [ontainerRelationshipDeleted](Visio.Pages.ContainerRelationshipDeleted.md)
- [onvertToGroupCanceled](Visio.Pages.ConvertToGroupCanceled.md)
- [ormulaChanged](Visio.Pages.FormulaChanged.md)
- [roupCanceled](Visio.Pages.GroupCanceled.md)
- [ageAdded](Visio.Pages.PageAdded.md)
- [ageChanged](Visio.Pages.PageChanged.md)
- [ageDeleteCanceled](Visio.Pages.PageDeleteCanceled.md)
- [ueryCancelConvertToGroup](Visio.Pages.QueryCancelConvertToGroup.md)
- [ueryCancelGroup](Visio.Pages.QueryCancelGroup.md)
- [ueryCancelPageDelete](Visio.Pages.QueryCancelPageDelete.md)
- [ueryCancelReplaceShapes](Visio.pages.querycancelreplaceshapes.md)
- [ueryCancelSelectionDelete](Visio.Pages.QueryCancelSelectionDelete.md)
- [ueryCancelUngroup](Visio.Pages.QueryCancelUngroup.md)
- [eplaceShapesCanceled](Visio.pages.replaceshapescanceled.md)
- [electionAdded](Visio.Pages.SelectionAdded.md)
- [electionDeleteCanceled](Visio.Pages.SelectionDeleteCanceled.md)
- [hapeAdded](Visio.Pages.ShapeAdded.md)
- [hapeChanged](Visio.Pages.ShapeChanged.md)
- [hapeDataGraphicChanged](Visio.Pages.ShapeDataGraphicChanged.md)
- [hapeExitedTextEdit](Visio.Pages.ShapeExitedTextEdit.md)
- [hapeLinkAdded](Visio.Pages.ShapeLinkAdded.md)
- [hapeLinkDeleted](Visio.Pages.ShapeLinkDeleted.md)
- [hapeParentChanged](Visio.Pages.ShapeParentChanged.md)
- [extChanged](Visio.Pages.TextChanged.md)
- [ngroupCanceled](Visio.Pages.UngroupCanceled.md)

## Methods

- [dd](Visio.Pages.Add.md)
- [etNames](Visio.Pages.GetNames.md)
- [etNamesU](Visio.Pages.GetNamesU.md)

## Properties

- [pplication](Visio.Pages.Application.md)
- [ount](Visio.Pages.Count.md)
- [ocument](Visio.Pages.Document.md)
- [ventList](Visio.Pages.EventList.md)
- [tem](Visio.Pages.Item.md)
- [temFromID](Visio.Pages.ItemFromID.md)
- [temU](Visio.Pages.ItemU.md)
- [bjectType](Visio.Pages.ObjectType.md)
- [ersistsEvents](Visio.Pages.PersistsEvents.md)
- [tat](Visio.Pages.Stat.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]