---
title: Cell object (Visio)
keywords: vis_sdr.chm10045
f1_keywords:
- vis_sdr.chm10045
ms.prod: visio
api_name:
- Visio.Cell
ms.assetid: 06ac28a6-5749-6c70-94bf-c721e217f375
ms.date: 06/19/2019
localization_priority: Normal
---


# Cell object (Visio)

Holds a formula that evaluates to some value.


## Remarks

The default property of a **Cell** object is **ResultIU**.

You can get or set a cell's formula or value. A cell belongs to a **Shape**, **Style**, or **Row** object and represents a property of the shape, style, or row. For example, the height of a shape equals the value of the shape's Height cell.

A program can control a shape's appearance and behavior by working with the formulas in the shape's cells. You can visually inspect most of a shape's cells by opening the shape's ShapeSheet window. 

Use the **[Cells](visio.shape.cells.md)** or **[CellsSRC](visio.shape.cellssrc.md)** property of a **Shape** object to retrieve a **Cell** object. To retrieve a cell in a style, use the **[Cells](visio.style.cells.md)** property of a **Style** object.


## Events

- [CellChanged](Visio.Cell.CellChanged.md)
- [FormulaChanged](Visio.Cell.FormulaChanged.md)

## Methods

- [GlueTo](Visio.Cell.GlueTo.md)
- [GlueToPos](Visio.Cell.GlueToPos.md)
- [Trigger](Visio.Cell.Trigger.md)

## Properties

- [Application](Visio.Cell.Application.md)
- [Column](Visio.Cell.Column.md)
- [ContainingMasterID](Visio.Cell.ContainingMasterID.md)
- [ContainingPageID](Visio.Cell.ContainingPageID.md)
- [ContainingRow](Visio.Cell.ContainingRow.md)
- [Dependents](Visio.Cell.Dependents.md)
- [Document](Visio.Cell.Document.md)
- [Error](Visio.Cell.Error.md)
- [EventList](Visio.Cell.EventList.md)
- [Formula](Visio.Cell.Formula.md)
- [FormulaForce](Visio.Cell.FormulaForce.md)
- [FormulaForceU](Visio.Cell.FormulaForceU.md)
- [FormulaU](Visio.Cell.FormulaU.md)
- [InheritedFormulaSource](Visio.Cell.InheritedFormulaSource.md)
- [InheritedValueSource](Visio.Cell.InheritedValueSource.md)
- [IsConstant](Visio.Cell.IsConstant.md)
- [IsInherited](Visio.Cell.IsInherited.md)
- [LocalName](Visio.Cell.LocalName.md)
- [Name](Visio.Cell.Name.md)
- [ObjectType](Visio.Cell.ObjectType.md)
- [PersistsEvents](Visio.Cell.PersistsEvents.md)
- [Precedents](Visio.Cell.Precedents.md)
- [Result](Visio.Cell.Result.md)
- [ResultForce](Visio.Cell.ResultForce.md)
- [ResultFromInt](Visio.Cell.ResultFromInt.md)
- [ResultFromIntForce](Visio.Cell.ResultFromIntForce.md)
- [ResultInt](Visio.Cell.ResultInt.md)
- [ResultIU](Visio.Cell.ResultIU.md)
- [ResultIUForce](Visio.Cell.ResultIUForce.md)
- [ResultStr](Visio.Cell.ResultStr.md)
- [ResultStrU](Visio.Cell.ResultStrU.md)
- [Row](Visio.Cell.Row.md)
- [RowName](Visio.Cell.RowName.md)
- [RowNameU](Visio.Cell.RowNameU.md)
- [Section](Visio.Cell.Section.md)
- [Shape](Visio.Cell.Shape.md)
- [Stat](Visio.Cell.Stat.md)
- [Style](Visio.Cell.Style.md)
- [Units](Visio.Cell.Units.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]