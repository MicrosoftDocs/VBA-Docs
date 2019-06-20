---
title: Page object (Visio)
keywords: vis_sdr.chm10190
f1_keywords:
- vis_sdr.chm10190
ms.prod: visio
api_name:
- Visio.Page
ms.assetid: 7a7f37ab-b448-eb70-b4f1-c185dfbd511e
ms.date: 05/08/2019
localization_priority: Normal
---


# Page object (Visio)

Represents a drawing page, which can be either a foreground page or a background page.


## Remarks

The default property of a **Page** object is **Name**.

To retrieve the active page in an instance, use the **[ActivePage](visio.application.activepage.md)** property of an **Application** object.

The members of a **Document** object's **[Pages](visio.pages.md)** collection represent the pages in that document. To retrieve a page's shapes, use the **Shapes** property of a **Page** object.


## Events

-  [AfterReplaceShapes](Visio.page.afterreplaceshapes.md)
-  [BeforePageDelete](Visio.Page.BeforePageDelete.md)
-  [BeforeReplaceShapes](Visio.page.beforereplaceshapes.md)
-  [BeforeSelectionDelete](Visio.Page.BeforeSelectionDelete.md)
-  [BeforeShapeDelete](Visio.Page.BeforeShapeDelete.md)
-  [BeforeShapeTextEdit](Visio.Page.BeforeShapeTextEdit.md)
-  [CalloutRelationshipAdded](Visio.Page.CalloutRelationshipAdded.md)
-  [CalloutRelationshipDeleted](Visio.Page.CalloutRelationshipDeleted.md)
-  [CellChanged](Visio.Page.CellChanged.md)
-  [ConnectionsAdded](Visio.Page.ConnectionsAdded.md)
-  [ConnectionsDeleted](Visio.Page.ConnectionsDeleted.md)
-  [ContainerRelationshipAdded](Visio.Page.ContainerRelationshipAdded.md)
-  [ContainerRelationshipDeleted](Visio.Page.ContainerRelationshipDeleted.md)
-  [ConvertToGroupCanceled](Visio.Page.ConvertToGroupCanceled.md)
-  [FormulaChanged](Visio.Page.FormulaChanged.md)
-  [GroupCanceled](Visio.Page.GroupCanceled.md)
-  [PageChanged](Visio.Page.PageChanged.md)
-  [PageDeleteCanceled](Visio.Page.PageDeleteCanceled.md)
-  [QueryCancelConvertToGroup](Visio.Page.QueryCancelConvertToGroup.md)
-  [QueryCancelGroup](Visio.Page.QueryCancelGroup.md)
-  [QueryCancelPageDelete](Visio.Page.QueryCancelPageDelete.md)
-  [QueryCancelReplaceShapes](Visio.page.querycancelreplaceshapes.md)
-  [QueryCancelSelectionDelete](Visio.Page.QueryCancelSelectionDelete.md)
-  [QueryCancelUngroup](Visio.Page.QueryCancelUngroup.md)
-  [ReplaceShapesCanceled](Visio.page.replaceshapescanceled.md)
-  [SelectionAdded](Visio.Page.SelectionAdded.md)
-  [SelectionDeleteCanceled](Visio.Page.SelectionDeleteCanceled.md)
-  [ShapeAdded](Visio.Page.ShapeAdded.md)
-  [ShapeChanged](Visio.Page.ShapeChanged.md)
-  [ShapeDataGraphicChanged](Visio.Page.ShapeDataGraphicChanged.md)
-  [ShapeExitedTextEdit](Visio.Page.ShapeExitedTextEdit.md)
-  [ShapeLinkAdded](Visio.Page.ShapeLinkAdded.md)
-  [ShapeLinkDeleted](Visio.Page.ShapeLinkDeleted.md)
-  [ShapeParentChanged](Visio.Page.ShapeParentChanged.md)
-  [TextChanged](Visio.Page.TextChanged.md)
-  [UngroupCanceled](Visio.Page.UngroupCanceled.md)


## Methods

-  [AddGuide](Visio.Page.AddGuide.md)
-  [AutoConnectMany](Visio.Page.AutoConnectMany.md)
-  [AutoSizeDrawing](Visio.Page.AutoSizeDrawing.md)
-  [AvoidPageBreaks](Visio.Page.AvoidPageBreaks.md)
-  [BoundingBox](Visio.Page.BoundingBox.md)
-  [CenterDrawing](Visio.Page.CenterDrawing.md)
-  [CreateSelection](Visio.Page.CreateSelection.md)
-  [Delete](Visio.Page.Delete.md)
-  [DrawArcByThreePoints](Visio.Page.DrawArcByThreePoints.md)
-  [DrawBezier](Visio.Page.DrawBezier.md)
-  [DrawCircularArc](Visio.Page.DrawCircularArc.md)
-  [DrawLine](Visio.Page.DrawLine.md)
-  [DrawNURBS](Visio.Page.DrawNURBS.md)
-  [DrawOval](Visio.Page.DrawOval.md)
-  [DrawPolyline](Visio.Page.DrawPolyline.md)
-  [DrawQuarterArc](Visio.Page.DrawQuarterArc.md)
-  [DrawRectangle](Visio.Page.DrawRectangle.md)
-  [DrawSpline](Visio.Page.DrawSpline.md)
-  [Drop](Visio.Page.Drop.md)
-  [DropCallout](Visio.Page.DropCallout.md)
-  [DropConnected](Visio.Page.DropConnected.md)
-  [DropContainer](Visio.Page.DropContainer.md)
-  [DropIntoList](Visio.Page.DropIntoList.md)
-  [DropLegend](Visio.Page.DropLegend.md)
-  [DropLinked](Visio.Page.DropLinked.md)
-  [DropMany](Visio.Page.DropMany.md)
-  [DropManyLinkedU](Visio.Page.DropManyLinkedU.md)
-  [DropManyU](Visio.Page.DropManyU.md)
-  [Duplicate](Visio.page.duplicate.md)
-  [Export](Visio.Page.Export.md)
-  [GetCallouts](Visio.Page.GetCallouts.md)
-  [GetContainers](Visio.Page.GetContainers.md)
-  [GetFormulas](Visio.Page.GetFormulas.md)
-  [GetFormulasU](Visio.Page.GetFormulasU.md)
-  [GetResults](Visio.Page.GetResults.md)
-  [GetShapesLinkedToData](Visio.Page.GetShapesLinkedToData.md)
-  [GetShapesLinkedToDataRow](Visio.Page.GetShapesLinkedToDataRow.md)
-  [GetTheme](Visio.page.gettheme.md)
-  [GetThemeVariant](Visio.page.getthemevariant.md)
-  [Import](Visio.Page.Import.md)
-  [InsertFromFile](Visio.Page.InsertFromFile.md)
-  [InsertObject](Visio.Page.InsertObject.md)
-  [Layout](Visio.Page.Layout.md)
-  [LayoutChangeDirection](Visio.Page.LayoutChangeDirection.md)
-  [LayoutIncremental](Visio.Page.LayoutIncremental.md)
-  [LinkShapesToDataRows](Visio.Page.LinkShapesToDataRows.md)
-  [OpenDrawWindow](Visio.Page.OpenDrawWindow.md)
-  [Paste](Visio.Page.Paste.md)
-  [PasteSpecial](Visio.Page.PasteSpecial.md)
-  [PasteToLocation](Visio.Page.PasteToLocation.md)
-  [Print](Visio.Page.Print.md)
-  [PrintTile](Visio.Page.PrintTile.md)
-  [ResizeToFitContents](Visio.Page.ResizeToFitContents.md)
-  [SetFormulas](Visio.Page.SetFormulas.md)
-  [SetResults](Visio.Page.SetResults.md)
-  [SetTheme](Visio.page.settheme.md)
-  [SetThemeVariant](Visio.page.setthemevariant.md)
-  [ShapeIDsToUniqueIDs](Visio.Page.ShapeIDsToUniqueIDs.md)
-  [SplitConnector](Visio.Page.SplitConnector.md)
-  [UniqueIDsToShapeIDs](Visio.Page.UniqueIDsToShapeIDs.md)
-  [VisualBoundingBox](Visio.page.visualboundingbox.md)


## Properties

-  [AlternativeText](Visio.Page.AlternativeText.md)
-  [Application](Visio.Page.Application.md)
-  [AutoSize](Visio.page.autosize.md)
-  [Background](Visio.Page.Background.md)
-  [BackPage](Visio.Page.BackPage.md)
-  [Comments](Visio.page.comments.md)
-  [Connects](Visio.Page.Connects.md)
-  [Document](Visio.Page.Document.md)
-  [EventList](Visio.Page.EventList.md)
-  [ID](Visio.Page.ID.md)
-  [Index](Visio.Page.Index.md)
-  [Layers](Visio.Page.Layers.md)
-  [LayoutRoutePassive](Visio.Page.LayoutRoutePassive.md)
-  [Name](Visio.Page.Name.md)
-  [NameU](Visio.Page.NameU.md)
-  [ObjectType](Visio.Page.ObjectType.md)
-  [OLEObjects](Visio.Page.OLEObjects.md)
-  [OriginalPage](Visio.Page.OriginalPage.md)
-  [PageSheet](Visio.Page.PageSheet.md)
-  [PersistsEvents](Visio.Page.PersistsEvents.md)
-  [Picture](Visio.Page.Picture.md)
-  [PrintTileCount](Visio.Page.PrintTileCount.md)
-  [ReviewerID](Visio.Page.ReviewerID.md)
-  [ShapeComments](Visio.page.shapecomments.md)
-  [Shapes](Visio.Page.Shapes.md)
-  [SpatialSearch](Visio.Page.SpatialSearch.md)
-  [Stat](Visio.Page.Stat.md)
-  [ThemeColors](Visio.Page.ThemeColors.md)
-  [ThemeEffects](Visio.Page.ThemeEffects.md)
-  [Title](Visio.Page.Title.md)
-  [Type](Visio.Page.Type.md)


## See also

- [Visio Object Model Reference](overview/visio/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]