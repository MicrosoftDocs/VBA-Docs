---
title: Chart.Export method (Project)
ms.prod: project-server
ms.assetid: 4f0ed821-f1c1-0e0b-0583-51b660ffad90
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Export method (Project)
Exports a chart in a graphic file format.

## Syntax

_expression_.**Export** (_bstr_, _varFilterName_, _varInteractive_)

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _bstr_|Required|**String**|The path and name of the exported file.|
| _varFilterName_|Optional|**Variant**|The language-independent name of the graphic filter as it appears in the registry ( `HKLM\\SOFTWARE\Wow6432Node\Microsoft\Shared Tools\Graphics Filters`).|
| _varInteractive_|Optional|**Variant**|**True** to display the dialog box that contains the filter-specific options, if any. If _varInteractive_ is **False**, Project uses the default values for the filter. The default value is **False**.|
| _bstr_|Required|**String**||
| _varFilterName_|Optional|**Variant**||
| _varInteractive_|Optional|**Variant**||

## Return value

 **Boolean**


## Remarks

The **Export** method overwrites an existing read/write file of the same name.


## Example

The following example exports the chart as a Portable Network Graphics (.png) file.


```vb
Sub ExportChart()
    Dim chartShape As Shape
    Dim reportName As String
    Dim fileFormat As String
    Dim fileName As String
    
    fileFormat = "PNG"
    fileName = "C:\Project\VBA\Samples\SimpleChart.png"
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    If (chartShape.Chart.Export(bstr:=filename, varFilterName:=fileFormat)) Then
        Debug.Print "Exported chart: " & filename
    End If
End Sub
```


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]