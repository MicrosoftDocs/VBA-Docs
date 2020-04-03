---
title: Shape object (Visio)
keywords: vis_sdr.chm10225
f1_keywords:
- vis_sdr.chm10225
ms.prod: visio
api_name:
- Visio.Shape
ms.assetid: da7a8872-4ebb-a607-e0ed-eebf68ff5630
ms.date: 05/08/2019
localization_priority: Normal
---


# Shape object (Visio)

Represents anything you can select in a drawing window: a basic shape, a group, a guide, or an object from another application embedded or linked in Microsoft Visio.


## Remarks

The default property of a **Shape** object is **Name**.

You can retrieve a particular **Shape** object from the **[Shapes](visio.shapes.md)** collection of the following objects:

- **Page** object
- **Master** object
- **Shape** object that represents a group
    
To retrieve **Cell** objects and **Connect** objects, use the **Cells** and **Connects** properties of a **Shape** object, respectively.

> [!NOTE] 
> The **PageSheet** property of a **[Page](visio.page.md)** object and **[Master](visio.master.md)** object returns a **Shape** object whose **Type** property returns **visTypePage**. It has cells that specify properties such as drawing size and drawing scale. 
> 
> The **[DocumentSheet](visio.document.documentsheet.md)** property of a **Document** object also returns a **Shape** object whose **Type** property returns **visTypeDoc**. It has cells that specify properties of the document.


## Events

- [BeforeSelectionDelete](Visio.Shape.BeforeSelectionDelete.md)
- [BeforeShapeDelete](Visio.Shape.BeforeShapeDelete.md)
- [BeforeShapeTextEdit](Visio.Shape.BeforeShapeTextEdit.md)
- [CellChanged](Visio.Shape.CellChanged.md)
- [ConvertToGroupCanceled](Visio.Shape.ConvertToGroupCanceled.md)
- [FormulaChanged](Visio.Shape.FormulaChanged.md)
- [GroupCanceled](Visio.Shape.GroupCanceled.md)
- [QueryCancelConvertToGroup](Visio.Shape.QueryCancelConvertToGroup.md)
- [QueryCancelGroup](Visio.Shape.QueryCancelGroup.md)
- [QueryCancelSelectionDelete](Visio.Shape.QueryCancelSelectionDelete.md)
- [QueryCancelUngroup](Visio.Shape.QueryCancelUngroup.md)
- [SelectionAdded](Visio.Shape.SelectionAdded.md)
- [SelectionDeleteCanceled](Visio.Shape.SelectionDeleteCanceled.md)
- [ShapeAdded](Visio.Shape.ShapeAdded.md)
- [ShapeChanged](Visio.Shape.ShapeChanged.md)
- [ShapeDataGraphicChanged](Visio.Shape.ShapeDataGraphicChanged.md)
- [ShapeExitedTextEdit](Visio.Shape.ShapeExitedTextEdit.md)
- [ShapeLinkAdded](Visio.Shape.ShapeLinkAdded.md)
- [ShapeLinkDeleted](Visio.Shape.ShapeLinkDeleted.md)
- [ShapeParentChanged](Visio.Shape.ShapeParentChanged.md)
- [TextChanged](Visio.Shape.TextChanged.md)
- [UngroupCanceled](Visio.Shape.UngroupCanceled.md)

## Methods

- [AddGuide](Visio.Shape.AddGuide.md)
- [AddHyperlink](Visio.Shape.AddHyperlink.md)
- [AddNamedRow](Visio.Shape.AddNamedRow.md)
- [AddRow](Visio.Shape.AddRow.md)
- [AddRows](Visio.Shape.AddRows.md)
- [AddSection](Visio.Shape.AddSection.md)
- [AddToContainers](Visio.Shape.AddToContainers.md)
- [AutoConnect](Visio.Shape.AutoConnect.md)
- [BoundingBox](Visio.Shape.BoundingBox.md)
- [BreakLinkToData](Visio.Shape.BreakLinkToData.md)
- [BringForward](Visio.Shape.BringForward.md)
- [BringToFront](Visio.Shape.BringToFront.md)
- [CenterDrawing](Visio.Shape.CenterDrawing.md)
- [ChangePicture](Visio.shape.changepicture.md)
- [ConnectedShapes](Visio.Shape.ConnectedShapes.md)
- [ConvertToGroup](Visio.Shape.ConvertToGroup.md)
- [Copy](Visio.Shape.Copy.md)
- [CreateSelection](Visio.Shape.CreateSelection.md)
- [CreateSubProcess](Visio.Shape.CreateSubProcess.md)
- [Cut](Visio.Shape.Cut.md)
- [Delete](Visio.Shape.Delete.md)
- [DeleteEx](Visio.Shape.DeleteEx.md)
- [DeleteRow](Visio.Shape.DeleteRow.md)
- [DeleteSection](Visio.Shape.DeleteSection.md)
- [Disconnect](Visio.Shape.Disconnect.md)
- [DrawArcByThreePoints](Visio.Shape.DrawArcByThreePoints.md)
- [DrawBezier](Visio.Shape.DrawBezier.md)
- [DrawCircularArc](Visio.Shape.DrawCircularArc.md)
- [DrawLine](Visio.Shape.DrawLine.md)
- [DrawNURBS](Visio.Shape.DrawNURBS.md)
- [DrawOval](Visio.Shape.DrawOval.md)
- [DrawPolyline](Visio.Shape.DrawPolyline.md)
- [DrawQuarterArc](Visio.Shape.DrawQuarterArc.md)
- [DrawRectangle](Visio.Shape.DrawRectangle.md)
- [DrawSpline](Visio.Shape.DrawSpline.md)
- [Drop](Visio.Shape.Drop.md)
- [DropMany](Visio.Shape.DropMany.md)
- [DropManyU](Visio.Shape.DropManyU.md)
- [Duplicate](Visio.Shape.Duplicate.md)
- [Export](Visio.Shape.Export.md)
- [FitCurve](Visio.Shape.FitCurve.md)
- [FlipHorizontal](Visio.Shape.FlipHorizontal.md)
- [FlipVertical](Visio.Shape.FlipVertical.md)
- [GetCustomPropertiesLinkedToData](Visio.Shape.GetCustomPropertiesLinkedToData.md)
- [GetCustomPropertyLinkedColumn](Visio.Shape.GetCustomPropertyLinkedColumn.md)
- [GetFormulas](Visio.Shape.GetFormulas.md)
- [GetFormulasU](Visio.Shape.GetFormulasU.md)
- [GetLinkedDataRecordsetIDs](Visio.Shape.GetLinkedDataRecordsetIDs.md)
- [GetLinkedDataRow](Visio.Shape.GetLinkedDataRow.md)
- [GetResults](Visio.Shape.GetResults.md)
- [GluedShapes](Visio.Shape.GluedShapes.md)
- [Group](Visio.Shape.Group.md)
- [HasCategory](Visio.Shape.HasCategory.md)
- [HitTest](Visio.Shape.HitTest.md)
- [Import](Visio.Shape.Import.md)
- [InsertFromFile](Visio.Shape.InsertFromFile.md)
- [InsertObject](Visio.Shape.InsertObject.md)
- [IsCustomPropertyLinked](Visio.Shape.IsCustomPropertyLinked.md)
- [Layout](Visio.Shape.Layout.md)
- [LinkToData](Visio.Shape.LinkToData.md)
- [MoveToSubprocess](Visio.Shape.MoveToSubprocess.md)
- [Offset](Visio.Shape.Offset.md)
- [OpenDrawWindow](Visio.Shape.OpenDrawWindow.md)
- [OpenSheetWindow](Visio.Shape.OpenSheetWindow.md)
- [Paste](Visio.Shape.Paste.md)
- [PasteSpecial](Visio.Shape.PasteSpecial.md)
- [RemoveFromContainers](Visio.Shape.RemoveFromContainers.md)
- [ReplaceShape](Visio.shape.replaceshape.md)
- [Resize](Visio.Shape.Resize.md)
- [ReverseEnds](Visio.Shape.ReverseEnds.md)
- [Rotate90](Visio.Shape.Rotate90.md)
- [SendBackward](Visio.Shape.SendBackward.md)
- [SendToBack](Visio.Shape.SendToBack.md)
- [SetBegin](Visio.Shape.SetBegin.md)
- [SetCenter](Visio.Shape.SetCenter.md)
- [SetEnd](Visio.Shape.SetEnd.md)
- [SetFormulas](Visio.Shape.SetFormulas.md)
- [SetQuickStyle](Visio.shape.setquickstyle.md)
- [SetResults](Visio.Shape.SetResults.md)
- [SwapEnds](Visio.Shape.SwapEnds.md)
- [TransformXYFrom](Visio.Shape.TransformXYFrom.md)
- [TransformXYTo](Visio.Shape.TransformXYTo.md)
- [Ungroup](Visio.Shape.Ungroup.md)
- [UpdateAlignmentBox](Visio.Shape.UpdateAlignmentBox.md)
- [VisualBoundingBox](Visio.shape.visualboundingbox.md)
- [XYFromPage](Visio.Shape.XYFromPage.md)
- [XYToPage](Visio.Shape.XYToPage.md)

## Properties

- [AlternativeText](Visio.Shape.AlternativeText.md)
- [Application](Visio.Shape.Application.md)
- [AreaIU](Visio.Shape.AreaIU.md)
- [CalloutsAssociated](Visio.Shape.CalloutsAssociated.md)
- [CalloutTarget](Visio.Shape.CalloutTarget.md)
- [CellExists](Visio.Shape.CellExists.md)
- [CellExistsU](Visio.Shape.CellExistsU.md)
- [Cells](Visio.Shape.Cells.md)
- [CellsRowIndex](Visio.Shape.CellsRowIndex.md)
- [CellsRowIndexU](Visio.Shape.CellsRowIndexU.md)
- [CellsSRC](Visio.Shape.CellsSRC.md)
- [CellsSRCExists](Visio.Shape.CellsSRCExists.md)
- [CellsU](Visio.Shape.CellsU.md)
- [Characters](Visio.Shape.Characters.md)
- [CharCount](Visio.Shape.CharCount.md)
- [ClassID](Visio.Shape.ClassID.md)
- [Comments](Visio.shape.comments.md)
- [Connects](Visio.Shape.Connects.md)
- [ContainerProperties](Visio.Shape.ContainerProperties.md)
- [ContainingMaster](Visio.Shape.ContainingMaster.md)
- [ContainingMasterID](Visio.Shape.ContainingMasterID.md)
- [ContainingPage](Visio.Shape.ContainingPage.md)
- [ContainingPageID](Visio.Shape.ContainingPageID.md)
- [ContainingShape](Visio.Shape.ContainingShape.md)
- [Data1](Visio.Shape.Data1.md)
- [Data2](Visio.Shape.Data2.md)
- [Data3](Visio.Shape.Data3.md)
- [DataGraphic](Visio.Shape.DataGraphic.md)
- [DistanceFrom](Visio.Shape.DistanceFrom.md)
- [DistanceFromPoint](Visio.Shape.DistanceFromPoint.md)
- [Document](Visio.Shape.Document.md)
- [EventList](Visio.Shape.EventList.md)
- [FillStyle](Visio.Shape.FillStyle.md)
- [FillStyleKeepFmt](Visio.Shape.FillStyleKeepFmt.md)
- [ForeignData](Visio.Shape.ForeignData.md)
- [ForeignType](Visio.Shape.ForeignType.md)
- [FromConnects](Visio.Shape.FromConnects.md)
- [GeometryCount](Visio.Shape.GeometryCount.md)
- [Help](Visio.Shape.Help.md)
- [Hyperlinks](Visio.Shape.Hyperlinks.md)
- [ID](Visio.Shape.ID.md)
- [Index](Visio.Shape.Index.md)
- [IsCallout](Visio.Shape.IsCallout.md)
- [IsDataGraphicCallout](Visio.Shape.IsDataGraphicCallout.md)
- [IsOpenForTextEdit](Visio.Shape.IsOpenForTextEdit.md)
- [Language](Visio.Shape.Language.md)
- [Layer](Visio.Shape.Layer.md)
- [LayerCount](Visio.Shape.LayerCount.md)
- [LengthIU](Visio.Shape.LengthIU.md)
- [LineStyle](Visio.Shape.LineStyle.md)
- [LineStyleKeepFmt](Visio.Shape.LineStyleKeepFmt.md)
- [Master](Visio.Shape.Master.md)
- [MasterShape](Visio.Shape.MasterShape.md)
- [MemberOfContainers](Visio.Shape.MemberOfContainers.md)
- [Name](Visio.Shape.Name.md)
- [NameID](Visio.Shape.NameID.md)
- [NameU](Visio.Shape.NameU.md)
- [Object](Visio.Shape.Object.md)
- [ObjectIsInherited](Visio.Shape.ObjectIsInherited.md)
- [ObjectType](Visio.Shape.ObjectType.md)
- [OneD](Visio.Shape.OneD.md)
- [Parent](Visio.Shape.Parent.md)
- [Paths](Visio.Shape.Paths.md)
- [PathsLocal](Visio.Shape.PathsLocal.md)
- [PersistsEvents](Visio.Shape.PersistsEvents.md)
- [Picture](Visio.Shape.Picture.md)
- [ProgID](Visio.Shape.ProgID.md)
- [RootShape](Visio.Shape.RootShape.md)
- [RowCount](Visio.Shape.RowCount.md)
- [RowExists](Visio.Shape.RowExists.md)
- [RowsCellCount](Visio.Shape.RowsCellCount.md)
- [RowType](Visio.Shape.RowType.md)
- [Section](Visio.Shape.Section.md)
- [SectionExists](Visio.Shape.SectionExists.md)
- [Shapes](Visio.Shape.Shapes.md)
- [SpatialNeighbors](Visio.Shape.SpatialNeighbors.md)
- [SpatialRelation](Visio.Shape.SpatialRelation.md)
- [SpatialSearch](Visio.Shape.SpatialSearch.md)
- [Stat](Visio.Shape.Stat.md)
- [Style](Visio.Shape.Style.md)
- [StyleKeepFmt](Visio.Shape.StyleKeepFmt.md)
- [Text](Visio.Shape.Text.md)
- [TextStyle](Visio.Shape.TextStyle.md)
- [TextStyleKeepFmt](Visio.Shape.TextStyleKeepFmt.md)
- [Title](Visio.Shape.Title.md)
- [Type](Visio.Shape.Type.md)
- [UniqueID](Visio.Shape.UniqueID.md)



## See also

- [Visio Object Model Reference](overview/visio/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
