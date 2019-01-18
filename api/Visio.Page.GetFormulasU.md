---
title: Page.GetFormulasU Method (Visio)
keywords: vis_sdr.chm10951935
f1_keywords:
- vis_sdr.chm10951935
ms.prod: visio
api_name:
- Visio.Page.GetFormulasU
ms.assetid: 8d7ba7d3-51e6-cd65-78ad-27640188e348
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.GetFormulasU Method (Visio)

Returns the formulas of many cells.


## Syntax

 _expression_. `GetFormulasU`( `_SID_SRCStream()_` , `_formulaArray()_` )

 _expression_ A variable that represents a [Page](./Visio.Page.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SID_SRCStream()_|Required| **Integer**|Stream identifying cells to be queried.|
| _formulaArray()_|Required| **Variant**|Out parameter. Array that receives formulas of queried cells.|

## Return value

Nothing


## Remarks

The  **GetFormulasU** method is like the **FormulaU** property of a **Cell** object, except you can use it to obtain the formulas of many cells at once, rather than one cell at a time. The **GetFormulasU** method is a specialization of the **GetResults** method, which can be used to obtain cell formulas or results. Setting up a call to the **GetFormulasU** method involves slightly less work than setting up the **GetResults** method.

You can use the  **GetFormulasU** method to get formulas in universal syntax of any set of cells in any set of shapes of the page or master.

 _SID_SRCStream()_ is an array of 2-byte integers. _SID_SRCStream()_ should be a one-dimensional array of 4 _n_ 2-byte integers for _n_ >= 1. The **GetFormulasU** method interprets _SID_SRCStream()_ as:




```vb
{sheetID, sectionIdx, rowIdx, cellIdx}n
```

 where _sheetID_ is the **ID** property of the **Shape** object on the page or master whose cell formula is desired.




 **Note**   If the _sheetID_ in an entry is **visInvalShapeID** (-1) or if the bottom byte of _sectionIdx_ is **visSectionInval** (255), the entry will be ignored and an empty variant will be returned in the corresponding results array entry. The reason for this is that the same _SID_SRCStream()_ array can be used on several calls to **GetFormulasU** , **SetFormulas** , and similar methods with the caller only needing to make minor changes to the stream between calls.

If the  **GetFormulasU** method succeeds, _formulaArray()_ returns a one-dimensional array of _n_ variants indexed from 0 to _n_ - 1. Each variant returns a formula as a string. _formulaArray()_ is an out parameter that is allocated by the **GetFormulasU** method, which passes ownership back to the caller. The caller should eventually perform the **SafeArrayDestroy** procedure on the returned array. Note that the **SafeArrayDestroy** procedure has the side effect of clearing the variants referenced by the array's entries, hence deallocating any strings the **GetFormulasU** method returns. (Microsoft Visual Basic and Microsoft Visual Basic for Applications take care of this for you.) The **GetFormulasU** method fails if _formulaArray()_ is **Null**.


 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

 As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the **GetFormulas** method to get more than one formula when you are using local syntax. Use the **GetFormulasU** method to get more than one formula when you are using universal syntax.


## Example

The following macro shows how to use the  **GetFormulasU** method. It assumes that there is an active Visio page that has at least three shapes on it. It uses **GetFormulasU** to get the width of shape 1, the height of shape 2, and the angle of shape 3.

This example uses the  **GetFormulasU** method of the **Page** object to get three cell formulas The input array has four slots for each cell, as it also would for **Master** objects. For **Shape** or **Style** objects, only three slots would be needed for each cell (section, row, and cell).




```vb
 
Public Sub GetFormulasU_Example() 
 
 On Error GoTo HandleError 
 
 Dim aintSheetSectionRowColumn(1 To 3 * 4) As Integer 
 
 aintSheetSectionRowColumn(1) = ActivePage.Shapes(1).ID 
 aintSheetSectionRowColumn(2) = visSectionObject 
 aintSheetSectionRowColumn(3) = visRowXFormOut 
 aintSheetSectionRowColumn(4) = visXFormWidth 
 
 aintSheetSectionRowColumn(5) = ActivePage.Shapes(2).ID 
 aintSheetSectionRowColumn(6) = visSectionObject 
 aintSheetSectionRowColumn(7) = visRowXFormOut 
 aintSheetSectionRowColumn(8) = visXFormHeight 
 
 aintSheetSectionRowColumn(9) = ActivePage.Shapes(3).ID 
 aintSheetSectionRowColumn(10) = visSectionObject 
 aintSheetSectionRowColumn(11) = visRowXFormOut 
 aintSheetSectionRowColumn(12) = visXFormAngle 
 
 'Return the formulas of the cells. 
 Dim avarFormulaArray() As Variant 
 ActivePage.GetFormulasU aintSheetSectionRowColumn, avarFormulaArray 
 
 Debug.Print "Shape 1 width is "; avarFormulaArray(0) 
 Debug.Print "Shape 2 height is "; avarFormulaArray(1) 
 Debug.Print "Shape 3 angle is "; avarFormulaArray(2) 
 
 
Exit Sub 
 
HandleError: 
 MsgBox "Error" 
Exit Sub 
 
End Sub
```


