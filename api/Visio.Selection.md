---
title: Selection object (Visio)
keywords: vis_sdr.chm10220
f1_keywords:
- vis_sdr.chm10220
ms.prod: visio
api_name:
- Visio.Selection
ms.assetid: e5734140-6dbe-7de8-9695-1a22fb4ac628
ms.date: 06/19/2019
localization_priority: Normal
---


# Selection object (Visio)

Represents a subset of **[Shape](visio.shape.md)** objects for a page or master to which an operation can be applied.


## Remarks

To retrieve a **Selection** object that corresponds to the set of shapes selected in a window, use the **[Selection](visio.window.selection.md)** property of a **Window** object.

The default property of a **Selection** object is **Item**.

After you retrieve a **Selection** object, you can add or remove shapes by using the **Select** method.

By default, the items reported by a **Selection** object do not include subselected or superselected **Shape** objects. Use the **IterationMode** property to control whether subselected and superselected **Shape** objects are reported. You can determine whether an individual item is subselected or superselected by using the **ItemStatus** property.


## Methods

-  [AddToContainers](Visio.Selection.AddToContainers.md)
-  [AddToGroup](Visio.Selection.AddToGroup.md)
-  [Align](Visio.Selection.Align.md)
-  [AutomaticLink](Visio.Selection.AutomaticLink.md)
-  [AvoidPageBreaks](Visio.Selection.AvoidPageBreaks.md)
-  [BoundingBox](Visio.Selection.BoundingBox.md)
-  [BreakLinkToData](Visio.Selection.BreakLinkToData.md)
-  [BringForward](Visio.Selection.BringForward.md)
-  [BringToFront](Visio.Selection.BringToFront.md)
-  [Combine](Visio.Selection.Combine.md)
-  [ConnectShapes](Visio.Selection.ConnectShapes.md)
-  [ConvertToGroup](Visio.Selection.ConvertToGroup.md)
-  [Copy](Visio.Selection.Copy.md)
-  [Cut](Visio.Selection.Cut.md)
-  [Delete](Visio.Selection.Delete.md)
-  [DeleteEx](Visio.Selection.DeleteEx.md)
-  [DeselectAll](Visio.Selection.DeselectAll.md)
-  [Distribute](Visio.Selection.Distribute.md)
-  [DrawRegion](Visio.Selection.DrawRegion.md)
-  [Duplicate](Visio.Selection.Duplicate.md)
-  [Export](Visio.Selection.Export.md)
-  [FitCurve](Visio.Selection.FitCurve.md)
-  [Flip](Visio.Selection.Flip.md)
-  [FlipHorizontal](Visio.Selection.FlipHorizontal.md)
-  [FlipVertical](Visio.Selection.FlipVertical.md)
-  [Fragment](Visio.Selection.Fragment.md)
-  [GetCallouts](Visio.Selection.GetCallouts.md)
-  [GetContainers](Visio.Selection.GetContainers.md)
-  [GetIDs](Visio.Selection.GetIDs.md)
-  [Group](Visio.Selection.Group.md)
-  [Intersect](Visio.Selection.Intersect.md)
-  [Join](Visio.Selection.Join.md)
-  [Layout](Visio.Selection.Layout.md)
-  [LayoutChangeDirection](Visio.Selection.LayoutChangeDirection.md)
-  [LayoutIncremental](Visio.Selection.LayoutIncremental.md)
-  [LinkToData](Visio.Selection.LinkToData.md)
-  [MemberOfContainersIntersection](Visio.Selection.MemberOfContainersIntersection.md)
-  [MemberOfContainersUnion](Visio.Selection.MemberOfContainersUnion.md)
-  [Move](Visio.Selection.Move.md)
-  [MoveToSubprocess](Visio.Selection.MoveToSubprocess.md)
-  [Offset](Visio.Selection.Offset.md)
-  [RemoveFromContainers](Visio.Selection.RemoveFromContainers.md)
-  [RemoveFromGroup](Visio.Selection.RemoveFromGroup.md)
-  [ReplaceShape](Visio.selection.replaceshape.md)
-  [Resize](Visio.Selection.Resize.md)
-  [ReverseEnds](Visio.Selection.ReverseEnds.md)
-  [Rotate](Visio.Selection.Rotate.md)
-  [Rotate90](Visio.Selection.Rotate90.md)
-  [Select](Visio.Selection.Select.md)
-  [SelectAll](Visio.Selection.SelectAll.md)
-  [SendBackward](Visio.Selection.SendBackward.md)
-  [SendToBack](Visio.Selection.SendToBack.md)
-  [SetContainerFormat](Visio.Selection.SetContainerFormat.md)
-  [SetQuickStyle](Visio.selection.setquickstyle.md)
-  [Subtract](Visio.Selection.Subtract.md)
-  [SwapEnds](Visio.Selection.SwapEnds.md)
-  [Trim](Visio.Selection.Trim.md)
-  [Ungroup](Visio.Selection.Ungroup.md)
-  [Union](Visio.Selection.Union.md)
-  [UpdateAlignmentBox](Visio.Selection.UpdateAlignmentBox.md)
-  [VisualBoundingBox](Visio.selection.visualboundingbox.md)

## Properties

-  [Application](Visio.Selection.Application.md)
-  [ContainingMaster](Visio.Selection.ContainingMaster.md)
-  [ContainingMasterID](Visio.Selection.ContainingMasterID.md)
-  [ContainingPage](Visio.Selection.ContainingPage.md)
-  [ContainingPageID](Visio.Selection.ContainingPageID.md)
-  [ContainingShape](Visio.Selection.ContainingShape.md)
-  [Count](Visio.Selection.Count.md)
-  [DataGraphic](Visio.Selection.DataGraphic.md)
-  [Document](Visio.Selection.Document.md)
-  [EventList](Visio.Selection.EventList.md)
-  [FillStyle](Visio.Selection.FillStyle.md)
-  [FillStyleKeepFmt](Visio.Selection.FillStyleKeepFmt.md)
-  [Item](Visio.Selection.Item.md)
-  [ItemStatus](Visio.Selection.ItemStatus.md)
-  [IterationMode](Visio.Selection.IterationMode.md)
-  [LineStyle](Visio.Selection.LineStyle.md)
-  [LineStyleKeepFmt](Visio.Selection.LineStyleKeepFmt.md)
-  [ObjectType](Visio.Selection.ObjectType.md)
-  [PersistsEvents](Visio.Selection.PersistsEvents.md)
-  [Picture](Visio.Selection.Picture.md)
-  [PrimaryItem](Visio.Selection.PrimaryItem.md)
-  [SelectionForDragCopy](Visio.Selection.SelectionForDragCopy.md)
-  [Stat](Visio.Selection.Stat.md)
-  [Style](Visio.Selection.Style.md)
-  [StyleKeepFmt](Visio.Selection.StyleKeepFmt.md)
-  [TextStyle](Visio.Selection.TextStyle.md)
-  [TextStyleKeepFmt](Visio.Selection.TextStyleKeepFmt.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]