---
title: DataColumn.SetProperty method (Visio)
keywords: vis_sdr.chm16760405
f1_keywords:
- vis_sdr.chm16760405
ms.prod: visio
api_name:
- Visio.DataColumn.SetProperty
ms.assetid: 5851daa0-e2e0-7073-7e26-f0fc73586b9b
ms.date: 02/16/2019
localization_priority: Normal
---


# DataColumn.SetProperty method (Visio)

Sets the value of the specified data-column property.

> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**SetProperty** (_Property_, _Value_)

_expression_ An expression that returns a **[DataColumn](Visio.DataColumn.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Property_|Required| **VisDataColumnProperties**|The data-column property whose value you want to set. See **Remarks** for possible values.|
| _Value_|Required| **Variant**|The value to assign the data-column property. See **Remarks** for possible values.|

## Return value

Nothing

## Remarks

When you link shapes in a Microsoft Visio drawing to data in a data recordset, Visio maps columns in the data recordset to rows in the Shape Data section of the ShapeSheet spreadsheet, each of which corresponds to a shape-data item. 

> [!NOTE] 
> In some previous versions of Visio, shape data were called custom properties.

Data-column properties map data columns to certain cells in the Shape Data section of the ShapeSheet. For example, by passing the **SetProperty** method a new value for the **DisplayName** property, which is represented by the enumerated value **visDataColumnPropertyDisplayName**, you set the value of the Label cell in the Shape Data section of the ShapeSheet for a particular shape data item. 

In addition, setting that property sets the label of the shape data item in the **Shape Data** dialog box, as well as the name of the data column that is displayed in the External Data window in the Visio user interface. These settings correspond to those that you can set in the **Column Settings** dialog box in the Visio user interface (right-click in the External Data window and then click **Column Settings**), as well as those that you can make in the **Types and Units** dialog box for each column (click **Data Types** in the **Column Settings** dialog box).

Possible values for the _Property_ parameter are declared in **VisDataColumnProperties**, and are shown in the following table.

|Constant|Value|Description|
|:-----|:-----|:-----|
| **visDataColumnPropertyCalendar**|3|Calendar of the data-column property.|
| **visDataColumnPropertyCurrency**|5|Currency of the data-column property.|
| **visDataColumnPropertyDisplayName**|6|Display name of the data-column property in the UI.|
| **visDataColumnPropertyHyperlink**|8|Whether the data-column value becomes a hyperlink in the Visio UI when it is linked to a shape.|
| **visDataColumnPropertyLangID**|2|Language ID of the data-column property.|
| **visDataColumnPropertyType**|1|Data type of the data-column property.|
| **visDataColumnPropertyUnits**|4|Units of the data-column property.|
| **visDataColumnPropertyVisible**|7|Whether the data-column property is visible in the UI.|

<br/>

Possible values for the _Value_ parameter depend on the _Property_ parameter value. The following table shows valid data-column property values for each data-column property, depending on the data-column data type.

|Data Column property| Number |Date |Currency |Duration |String |Boolean |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|Constant|**visPropTypeNumber**| **visPropTypeDate**| **visPropTypeCurrency**| **visPropTypeDuration**| **visPropTypeString**| **visPropTypeBoolean**|
|Visible| **Boolean**| **Boolean**| **Boolean**| **Boolean**| **Boolean**| **Boolean**|
|DisplayName| **String**| **String**| **String**| **String**| **String**| **String**|
|LangID|Valid LCID number|||||
|Currency|||Valid 3-letter currency-constant string as used in the CY function in the Visio ShapeSheet spreadsheet.||||
|Calendar||One of the members of **VisCellVals**, depending on the LangID value (see table below).|||||
|Units|One of the following members of **VisUnitsCodes**:<ul><li><b>visAcre</b></li><li><b>visAngleUnits</b></li><li><b>visCentimeters</b></li><li><b>visCiceros</b></li><li><b>visCicerosAndDidots</b></li><li><b>visDegreeMinSec</b></li><li><b>visDegrees</b></li><li><b>visDrawingUnits</b></li><li><b>visFeet</b></li><li><b>visFeetAndInches</b></li><li><b>visHectare</b></li><li><b>visDidots</b></li><li><b>visInches</b></li><li><b>visInchFrac</b></li><li><b>visKilometers</b></li><li><b>visMeters</b></li><li><b>visMileFrac</b></li><li><b>visMiles</b></li><li><b>visMillimeters</b></li><li><b>visMin</b></li><li><b>visNautMiles</b></li><li><b>visPageUnits</b></li><li><b>visPicas</b></li><li><b>visPicasAndPoints</b></li><li><b>visPoints</b></li><li><b>visRadians</b></li><li><b>visSec</b></li><li><b>visYards</b></li><li><b>visNumber</b>  (special behavior: this constant makes the value unitless)  </li></ul><br/>OR<br/><br/>Descriptive string: a string used for units, such as _cm_ or _sq cm_. This string will be validated so that it is one of the supported Visio units. Passing invalid strings causes the method to fail.|||One of the following members of **VisUnitsCodes**:<ul><li><b>visDurationUnits</b></li><li><b>visElapsedDay</b></li><li><b>visElapsedHour</b></li><li><b>visElapsedMin</b></li><li><b>visElapsedSec</b></li><li><b>visElapsedWeek</b></li></ul><br/>OR<br/><br/>Descriptive string: a string used for units such as _ew_. This string will be validated so that it is one of the supported Visio units. Passing an invalid string will cause this method to fail.|||
|HyperLink||||| **Boolean**||

<br/>

The LangID and Calendar properties are bound by the validation rules shown in the following table. Languages not shown use the Western calendar only.

|Language|Hirji|Western|French Transliterated|English Transliterated|Hebrew Lunar|Saka Era|Japanese Emperor Era|Korean Danki|Thai Buddhist|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|All Arabic|x|x|x|x||||||
|Bengali(Bangladesh)|x|x||||||||
|Divehi|x|x||||||||
|All English|x|x|||x|x||||
|Persian|x|x||||||||
|Hebrew|x||||x|||||
|Hindi|x|||||x||||
|Japanese||x|||||x|||
|Korean||x||||||x||
|Kashmiri (Arabic)|x|x||||||||
|Punjabi (Pakistan)|x|x||||||||
|Pashto|x|x||||||||
|Sindhi|x|x||||||||
|Thai||||||||||
|Urdu|x|x||||||||
|Tamzight|x|x||||||||

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **GetProperty** method to get the value of the Label cell in the Shape Data section for the first column in the data recordset passed to the method, and display it in the Immediate window. It then uses the **SetProperty** method to set the value and displays the new value. Changing this value changes the label of the shape data item in the **Shape Data** dialog box for all shapes linked to rows in the data recordset.

To get and set the Label cell value, the macro passes the **visDataColumnPropertyDisplayName** value from the **VisDataColumnProperties** enumeration to the **DataColumn.GetProperty** and **DataColumn.SetProperty** methods.

Before running this macro, create at least one data recordset in your VBA project to pass to the macro.

```vb
 
Public Sub SetProperty_Example(vsoDataRecordset As Visio.DataRecordset) 
    Dim strPropertyName As String 
    Dim strNewName As String 
    Dim vsoDataColumn As Visio.DataColumn 
 
    strNewName = "New Property Name" 
    Set vsoDataColumn = vsoDataRecordset.DataColumns(1) 
 
    strPropertyName = vsoDataColumn.GetProperty(visDataColumnPropertyDisplayName) 
    Debug.Print strPropertyName 
 
    vsoDataColumn.SetProperty visDataColumnPropertyDisplayName, strNewName 
    strPropertyName = vsoDataColumn.GetProperty(visDataColumnPropertyDisplayName) 
    Debug.Print strPropertyName 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]