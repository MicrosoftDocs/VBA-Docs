---
title: Documents object (Visio)
keywords: vis_sdr.chm10085
f1_keywords:
- vis_sdr.chm10085
ms.prod: visio
api_name:
- Visio.Documents
ms.assetid: e9291149-964e-c6fb-4c62-bf2f35a6a0a7
ms.date: 06/19/2019
localization_priority: Normal
---


# Documents object (Visio)

Includes a **[Document](Visio.Document.md)** object for each open document in a Microsoft Visio instance.


## Remarks

To retrieve a **Documents** collection, use the **[Documents](visio.application.documents.md)** property of an **Application** object.

The default property of a **Documents** collection is **Item**.

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this collection maps to the following types:

- **Microsoft.Office.Interop.Visio.IVDocuments.GetEnumerator()** (to enumerate the **Document** objects)   
- **Microsoft.Office.Interop.Visio.IVDocuments** (to access the collection)

## Events

- [AfterDocumentMerge](Visio.documents.afterdocumentmerge.md)
- [AfterRemoveHiddenInformation](Visio.Documents.AfterRemoveHiddenInformation.md)
- [AfterReplaceShapes](Visio.documents.afterreplaceshapes.md)
- [BeforeDataRecordsetDelete](Visio.Documents.BeforeDataRecordsetDelete.md)
- [BeforeDocumentClose](Visio.Documents.BeforeDocumentClose.md)
- [BeforeDocumentSave](Visio.Documents.BeforeDocumentSave.md)
- [BeforeDocumentSaveAs](Visio.Documents.BeforeDocumentSaveAs.md)
- [BeforeMasterDelete](Visio.Documents.BeforeMasterDelete.md)
- [BeforePageDelete](Visio.Documents.BeforePageDelete.md)
- [BeforeReplaceShapes](Visio.documents.beforereplaceshapes.md)
- [BeforeSelectionDelete](Visio.Documents.BeforeSelectionDelete.md)
- [BeforeShapeDelete](Visio.Documents.BeforeShapeDelete.md)
- [BeforeShapeTextEdit](Visio.Documents.BeforeShapeTextEdit.md)
- [BeforeStyleDelete](Visio.Documents.BeforeStyleDelete.md)
- [CalloutRelationshipAdded](Visio.Documents.CalloutRelationshipAdded.md)
- [CalloutRelationshipDeleted](Visio.Documents.CalloutRelationshipDeleted.md)
- [CellChanged](Visio.Documents.CellChanged.md)
- [ConnectionsAdded](Visio.Documents.ConnectionsAdded.md)
- [ConnectionsDeleted](Visio.Documents.ConnectionsDeleted.md)
- [ContainerRelationshipAdded](Visio.Documents.ContainerRelationshipAdded.md)
- [ContainerRelationshipDeleted](Visio.Documents.ContainerRelationshipDeleted.md)
- [ConvertToGroupCanceled](Visio.Documents.ConvertToGroupCanceled.md)
- [DataRecordsetAdded](Visio.Documents.DataRecordsetAdded.md)
- [DataRecordsetChanged](Visio.Documents.DataRecordsetChanged.md)
- [DesignModeEntered](Visio.Documents.DesignModeEntered.md)
- [DocumentChanged](Visio.Documents.DocumentChanged.md)
- [DocumentCloseCanceled](Visio.Documents.DocumentCloseCanceled.md)
- [DocumentCreated](Visio.Documents.DocumentCreated.md)
- [DocumentOpened](Visio.Documents.DocumentOpened.md)
- [DocumentSaved](Visio.Documents.DocumentSaved.md)
- [DocumentSavedAs](Visio.Documents.DocumentSavedAs.md)
- [FormulaChanged](Visio.Documents.FormulaChanged.md)
- [GroupCanceled](Visio.Documents.GroupCanceled.md)
- [MasterAdded](Visio.Documents.MasterAdded.md)
- [MasterChanged](Visio.Documents.MasterChanged.md)
- [MasterDeleteCanceled](Visio.Documents.MasterDeleteCanceled.md)
- [PageAdded](Visio.Documents.PageAdded.md)
- [PageChanged](Visio.Documents.PageChanged.md)
- [PageDeleteCanceled](Visio.Documents.PageDeleteCanceled.md)
- [QueryCancelConvertToGroup](Visio.Documents.QueryCancelConvertToGroup.md)
- [QueryCancelDocumentClose](Visio.Documents.QueryCancelDocumentClose.md)
- [QueryCancelGroup](Visio.Documents.QueryCancelGroup.md)
- [QueryCancelMasterDelete](Visio.Documents.QueryCancelMasterDelete.md)
- [QueryCancelPageDelete](Visio.Documents.QueryCancelPageDelete.md)
- [QueryCancelReplaceShapes](Visio.documents.querycancelreplaceshapes.md)
- [QueryCancelSelectionDelete](Visio.Documents.QueryCancelSelectionDelete.md)
- [QueryCancelStyleDelete](Visio.Documents.QueryCancelStyleDelete.md)
- [QueryCancelUngroup](Visio.Documents.QueryCancelUngroup.md)
- [ReplaceShapesCanceled](Visio.documents.replaceshapescanceled.md)
- [RuleSetValidated](Visio.Documents.RuleSetValidated.md)
- [RunModeEntered](Visio.Documents.RunModeEntered.md)
- [SelectionAdded](Visio.Documents.SelectionAdded.md)
- [SelectionDeleteCanceled](Visio.Documents.SelectionDeleteCanceled.md)
- [ShapeAdded](Visio.Documents.ShapeAdded.md)
- [ShapeChanged](Visio.Documents.ShapeChanged.md)
- [ShapeDataGraphicChanged](Visio.Documents.ShapeDataGraphicChanged.md)
- [ShapeExitedTextEdit](Visio.Documents.ShapeExitedTextEdit.md)
- [ShapeLinkAdded](Visio.Documents.ShapeLinkAdded.md)
- [ShapeLinkDeleted](Visio.Documents.ShapeLinkDeleted.md)
- [ShapeParentChanged](Visio.Documents.ShapeParentChanged.md)
- [StyleAdded](Visio.Documents.StyleAdded.md)
- [StyleChanged](Visio.Documents.StyleChanged.md)
- [StyleDeleteCanceled](Visio.Documents.StyleDeleteCanceled.md)
- [TextChanged](Visio.Documents.TextChanged.md)
- [UngroupCanceled](Visio.Documents.UngroupCanceled.md)

## Methods

- [Add](Visio.Documents.Add.md)
- [AddEx](Visio.Documents.AddEx.md)
- [CanCheckOut](Visio.Documents.CanCheckOut.md)
- [CheckOut](Visio.Documents.CheckOut.md)
- [GetNames](Visio.Documents.GetNames.md)
- [Open](Visio.Documents.Open.md)
- [OpenEx](Visio.Documents.OpenEx.md)

## Properties

- [Application](Visio.Documents.Application.md)
- [Count](Visio.Documents.Count.md)
- [EventList](Visio.Documents.EventList.md)
- [Item](Visio.Documents.Item.md)
- [ItemFromID](Visio.Documents.ItemFromID.md)
- [ObjectType](Visio.Documents.ObjectType.md)
- [PersistsEvents](Visio.Documents.PersistsEvents.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]