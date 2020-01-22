---
title: What's new for VBA in Office 2019
ms.prod: office
ms.date: 01/22/2020
localization_priority: Normal
---

# What's new for VBA in Office 2019

The following tables summarize the new Visual Basic for Applications (VBA) language updates for Office 2019.

## Office

|Name|Description|
|:-----|:-----|
|**[MsoGraphicStyleIndex enumeration (Office)](../../api/office.msoshapestyleindex.md)**|Represents preset Graphic styles.|
|**[MsoShapeType enumeration (Office)](../../api/office.msoshapetype.md)** | **mso3DModel**, **msoGraphic**, **msoLinked3DModel**, and **msoLinkedGraphic** values are now included.|

## Access

|Name|Description|
|:-----|:-----|
|**[Chart object (Access)](../../api/access.chart.md)** | A customizable visualization of data that can be included in a report or a form.|
|**[ChartAxis object (Access)](../../api/access.chartaxis.md)** | Represents a field whose values will be included in the category axis of a chart.|
|**[ChartSeries object (Access)](../../api/access.chartseries.md)** | Represents a series of values in the specified chart.|
|**[ChartValues object (Access)](../../api/access.chartvalues.md)** | Represents a field whose values will be plotted in the value (Y) axis of a chart.|
|**[acCommand enumeration (Access)](../../api/access.accommand.md)** | **acCmdImportAttachdBase** and **acCmdExportdBase** values are now included.|

## Excel

|Name|Description|
|:-----|:-----|
|**[Model3DFormat object (Excel)](../../api/excel.model3dformat.md)** | Represents the properties of a 3D model shape.|
|**[Series.GeoMappingLevel property (Excel)](../../api/excel.series.geomappinglevel.md)**|Specifies the geography mapping level for the specified series within the chart group. Read/write **XlGeoMappingLevel**.|
|**[Series.GeoProjectionType property (Excel)](../../api/excel.series.geoprojectiontype.md)**|Specifies the geography projection type for the specified series within the chart group. Read/write **XlGeoProjectionType**.|
|**[Series.RegionLabelOptions property (Excel)](../../api/excel.series.regionlabeloptions.md)**|Specifies the region labelling behavior for the specified series within the chart group. Read/write **XlRegionLabelOptions**.|
|**[Shape.GraphicStyle property (Excel)](../../api/excel.shape.graphicstyle.md)**|Returns or sets an **MsoGraphicStyleIndex** that represents the style of an SVG graphic. Read/write.|
|**[Shape.Model3D property (Excel)](../../api/excel.shape.model3d.md)**|Returns a **Model3DFormat** object that represents the 3D properties of a 3D model object. Read-only.|
|**[ShapeRange.GraphicStyle property (Excel)](../../api/excel.shaperange.graphicstyle.md)**|Returns or sets an **MsoGraphicStyleIndex** that represents the style of a shape range containing one or more SVG graphics. Read/write.|
|**[ShapeRange.Model3D property (Excel)](../../api/excel.shaperange.model3d.md)**|Returns a **Model3DFormat** object that represents the 3D properties of a 3D model object. Read-only.|
|**[Shapes.Add3DModel method (Excel)](../../api/excel.shapes.add3dmodel.md)**|Creates a 3D model from an existing file. Returns a **Shape** object that represents the new 3D model.|
|**[Excel.XlGeoMappingLevel enumeration (Excel)](../../api/excel.xlgeomappinglevel.md)**|Constants passed to and returned by the **Series.GeoMappingLevel** property.|
|**[Excel.XlGeoProjectionType enumeration (Excel)](../../api/excel.xlgeoprojectiontype.md)**|Constants passed to and returned by the **Series.GeoProjectionType** property.|
|**[Excel.XlRegionLabelOptions enumeration (Excel)](../../api/excel.xlregionlabeloptions.md)**|Constants passed to and returned by the **Series.RegionLabelOptions** property.|

## PowerPoint

|Name|Description|
|:-----|:-----|
|**[Model3DFormat object (PowerPoint)](../../api/powerpoint.model3dformat.md)** | Represents the properties of a 3D model shape.|
|**[Presentation.AutoSaveOn property (PowerPoint)](../../api/powerpoint.presentation.autosaveon.md)** | **True** if the edits in the presentation are automatically saved. Read/write **Boolean**.|
|**[Shape.Decorative property (PowerPoint)](../../api/powerpoint.shape.decorative.md)**|Sets or returns the decorative flag for the specified object. Read/write.|
|**[Shape.GraphicStyle property (PowerPoint)](../../api/powerpoint.shape.graphicstyle.md)**|Returns or sets an **MsoGraphicStyleIndex** that represents the style of an SVG graphic. Read/write.|
|**[Shape.Model3D property (PowerPoint)](../../api/powerpoint.shape.model3d.md)**|Returns a **Model3DFormat** object that represents the 3D properties of a 3D model object. Read-only.|
|**[ShapeRange.Decorative property (PowerPoint)](../../api/powerpoint.shaperange.decorative.md)**|Sets or returns the decorative flag for the specified object. Read/write.|
|**[ShapeRange.GraphicStyle property (PowerPoint)](../../api/powerpoint.shaperange.graphicstyle.md)**|Returns or sets an **MsoGraphicStyleIndex** that represents the style of a shape range containing one or more SVG graphics. Read/write.|
|**[ShapeRange.Model3D property (PowerPoint)](../../api/powerpoint.shaperange.model3d.md)**|Returns a **Model3DFormat** object that represents the 3D properties of a 3D model object. Read-only.|
|**[Shapes.Add3DModel method (PowerPoint)](../../api/powerpoint.shapes.add3dmodel.md)**|Creates a **Model3DFormat** object from an existing file. Returns a **Shape** object that represents the new 3D model.|

## Visio

|Name|Description|
|:-----|:-----|
|**[Shape.AlternativeText property (Visio)](../../api/visio.shape.alternativetext.md)**|Returns or sets the alternative text description associated with an object. Read/write.|
|**[Shape.Title property (Visio)](../../api/visio.shape.title.md)**|Returns or sets the alternative text associated with an object. Read/write.|

## Word

|Name|Description|
|:-----|:-----|
|**[Model3DFormat object (Word)](../../api/word.model3dformat.md)** |Represents the properties of a 3D model shape.|
|**[Shape.GraphicStyle property (Word)](../../api/word.shape.graphicstyle.md)**|Returns or sets an **MsoGraphicStyleIndex** that represents the style of an SVG graphic. Read/write.|
|**[Shape.Model3D property (Word)](../../api/word.shape.model3d.md)**|Returns a **Model3DFormat** object that represents the 3D properties of a 3D model object. Read-only.|
|**[ShapeRange.GraphicStyle property (Word)](../../api/word.shaperange.graphicstyle.md)**|Returns or sets an **MsoGraphicStyleIndex** that represents the style of a shape range containing one or more SVG graphics. Read/write.|
|**[ShapeRange.Model3D property (Word)](../../api/word.shaperange.model3d.md)**|Returns a **Model3DFormat** object that represents the 3D properties of a 3D model object. Read-only.|
|**[Shapes.Add3DModel method (Word)](../../api/word.shapes.add3dmodel.md)**|Adds a 3D model to a drawing canvas. Returns a **Shape** object that represents the 3D model and adds it to the **CanvasShapes** collection.|
|**[WdInlineShapeType enumeration (Word)](../../api/word.wdinlineshapetype.md)** | **wdInlineShape3DModel** and **wdInlineShapeLinked3DModel** values are now included.|

## See also

- [Library reference VBA](../../api/overview/library-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
