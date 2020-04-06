---
title: Cell.Formula property (Visio)
keywords: vis_sdr.chm10113560
f1_keywords:
- vis_sdr.chm10113560
ms.prod: visio
api_name:
- Visio.Cell.Formula
ms.assetid: 36636047-9ee7-d461-92fb-0b36853e3201
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.Formula property (Visio)

Gets or sets the formula for a  **Cell** object. Read/write.


## Syntax

_expression_.**Formula**

_expression_ A variable that represents a **[Cell](Visio.Cell.md)** object.


## Return value

String


## Remarks

If a cell's formula is protected with the GUARD function, you must use the  **FormulaForce** property to change the cell's formula.




> [!NOTE] 
> Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **Formula** property to get a cell's formula string in local syntax or to use a mix of local and universal syntax to set it. Use the **FormulaU** property to get or parse a formula in universal syntax. When you use **FormulaU**, the decimal point is always ".", the delimiter is always ",", and you must use universal unit strings (for details on universal strings, see [About Units of Measure](../visio/Concepts/about-units-of-measure-visio.md)).


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Formula** property to set the formula for a ShapeSheet cell. It draws a rectangle on a page and bows, or curves the lines of the rectangle by changing the shape's lines to arcs. Then it draws an inner rectangle within the bowed lines of the first rectangle.


```vb
 
Public Sub Formula_Example() 
  
    Dim vsoPage As Visio.Page  
    Dim vsoShape As Visio.Shape  
    Dim vsoCell As Visio.Cell  
    Dim strBowCell As String 
    Dim strBowFormula As String 
    Dim intIndex As Integer 
    Dim intCounter As Integer 
 
    'Set the value of the strBowCell string.  
    strBowCell = "Scratch.X1"  
 
    'Set the value of the strBowFormula string.  
    strBowFormula = "=Min(Width, Height) / 5"  
 
    Set vsoPage = ActivePage  
 
    'If there isn't an active page, set vsoPage 
    'to the first page of the active document. 
    If vsoPage Is Nothing Then 
    Set vsoPage = ActiveDocument.Pages(1)  
    End If   
 
    'Draw a rectangle on the active page. 
    Set vsoShape = vsoPage.DrawRectangle(1, 5, 5, 1)  
 
    'Add a Scratch section to the shape's ShapeSheet.  
    vsoShape.AddSection visSectionScratch  
 
    'Add a row to the Scratch section.  
    vsoShape.AddRow visSectionScratch, visRowScratch, 0  
 
    'Set vsoCell to the Scratch.X1 cell and set its formula. 
    Set vsoCell = vsoShape.Cells(strBowCell)  
    vsoCell.Formula = strBowFormula  
 
    'Bow in or curve the rectangle's lines by changing 
    'each row type from LineTo to ArcTo and entering the bow value. 
    For intCounter = 1 To 4  
        vsoShape.RowType(visSectionFirstComponent, visRowVertex + intCounter) = visTagArcTo  
        Set vsoCell = vsoShape.CellsSRC(visSectionFirstComponent, visRowVertex + intCounter, 2)  
        vsoCell.Formula = "-" & strBowCell  
    Next intCounter  
 
    'Create an inner rectangle. 
    'Set the section index for the inner rectangle's Geometry section.  
    intIndex = visSectionFirstComponent + 1  
 
    'Add an inner rectangle Geometry section.  
    vsoShape.AddSection intIndex  
 
    'Add the first 2 rows to the section.  
    vsoShape.AddRow intIndex, visRowComponent, visTagComponent  
    vsoShape.AddRow intIndex, visRowVertex, visTagMoveTo  
 
    'Add 4 LineTo rows to the section. 
    For intCounter = 1 To 4  
        vsoShape.AddRow intIndex, visRowLast, visTagLineTo  
    Next intCounter  
 
    'Set the inner rectangle start point cell formulas. 
    Set vsoCell = vsoShape.CellsSRC(intIndex, 1, 0)  
    vsoCell.Formula = "Width * 0 + " & strBowCell  
    Set vsoCell = vsoShape.CellsSRC(intIndex, 1, 1)  
    vsoCell.Formula = "Height * 0 + " & strBowCell  
 
    'Draw the inner rectangle bottom line. 
    Set vsoCell = vsoShape.CellsSRC(intIndex, 2, 0)  
    vsoCell.Formula = "Width * 1 - " & strBowCell  
    Set vsoCell = vsoShape.CellsSRC(intIndex, 2, 1)  
    vsoCell.Formula = "Height * 0 + " & strBowCell  
 
    'Draw the inner rectangle right side line. 
    Set vsoCell = vsoShape.CellsSRC(intIndex, 3, 0)  
    vsoCell.Formula = "Width * 1 - " & strBowCell  
    Set vsoCell = vsoShape.CellsSRC(intIndex, 3, 1)  
    vsoCell.Formula = "Height * 1 - " & strBowCell  
 
    'Draw the inner rectangle top line. 
    Set vsoCell = vsoShape.CellsSRC(intIndex, 4, 0)  
    vsoCell.Formula = "Width * 0 + " & strBowCell  
    Set vsoCell = vsoShape.CellsSRC(intIndex, 4, 1)  
    vsoCell.Formula = "Height * 1 - " & strBowCell  
 
    'Draw the inner rectangle left side line. 
    Set vsoCell = vsoShape.CellsSRC(intIndex, 5, 0)  
    vsoCell.Formula = "Geometry2.X1"  
    Set vsoCell = vsoShape.CellsSRC(intIndex, 5, 1)  
    vsoCell.Formula = "Geometry2.Y1"  
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]