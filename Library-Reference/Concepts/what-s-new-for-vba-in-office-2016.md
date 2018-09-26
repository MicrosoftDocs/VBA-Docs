---
title: What's New for VBA in Office 2016
ms.prod: office
ms.assetid: c0294abb-bc0e-495d-b387-4398378dd3ad
ms.date: 06/08/2017
---


# What's New for VBA in Office 2016
The following tables summarize the new VBA language updates for Office 2016.




## Access





|**Name**|**Description**|
|:-----|:-----|
|**[CodeProject.IsSQLBackend property (Access)](../../api/overview/Library-Reference.md)**|Returns the  **Boolean** value **true** if the code project was created in Access 2013 and newer, and **false** if otherwise.|
|**[CurrentProject.IsSQLBackend property (Access)](../../api/overview/Library-Reference.md)**|Returns  **true** if the current project was created in Access 2013 and onwards and **false** if the current project was created prior to Access 2013. Read-only **Boolean**.|

## Excel





|**Name**|**Description**|
|:-----|:-----|
|**[Chart.ShowExpandCollapseEntireFieldButtons Property (Excel)](../../api/Excel.chart.showexpandcollapseentirefieldbuttons.md)**|**True** to display the **Expand Entire Field** and **Collapse Entire Field** buttons on the specified pivot chart. Read/write **Boolean**.|
|**[ChartGroup.BinsCountValue Property (Excel)](../../api/Excel.chartgroup.binscountvalue.md)**|Specifies the number of bins in the histogram chart. Read/write  **Long**.|
|**[ChartGroup.BinsOverflowEnabled Property (Excel)](../../api/Excel.chartgroup.binsoverflowenabled.md)**|Specifies whether a bin for values above the [BinsOverflowValue](../../api/Excel.chartgroup.binsoverflowvalue.md) is enabled. Read/write **Boolean**.|
|**[ChartGroup.BinsOverflowValue Property (Excel)](../../api/Excel.chartgroup.binsoverflowvalue.md)**|If an [BinsOverflowEnabled](../../api/Excel.chartgroup.binsoverflowenabled.md) is **True**, specifies the value above which an overflow bin is displayed. Read/write **Double**.|
|**[ChartGroup.BinsType Property (Excel)](../../api/Excel.chartgroup.binstype.md)**|Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Read/write [XlBinsType](../../api/Excel.xlbinstype.md).|
|**[ChartGroup.BinsUnderflowEnabled Property (Excel)](../../api/Excel.chartgroup.binsunderflowenabled.md)**|Specifies whether a bin for values below the [BinsUnderflowValue](../../api/Excel.chartgroup.binsunderflowvalue.md) is enabled. Read/write **Boolean**.|
|**[ChartGroup.BinsUnderflowValue Property (Excel)](../../api/Excel.chartgroup.binsunderflowvalue.md)**|If an [BinsUnderflowEnabled](../../api/Excel.chartgroup.binsunderflowenabled.md) is **True**, specifies the value below which an underflow bin is displayed. Read/write **Double**.|
|**[ChartGroup.BinWidthValue Property (Excel)](../../api/Excel.chartgroup.binwidthvalue.md)**|Specifies the number of points in each range. Read/write  **Double**.|
|**[CubeField.AutoGroup Method (Excel)](../../api/Excel.cubefield.autogroup.md)**|Automatically groups the cube fields in an OLAP cube, optionally in the specified orientation and/or at the specified position.|
|**[Model.ModelFormatBoolean Property (Excel)](../../api/Excel.model.modelformatboolean.md)**|Returns a [ModelFormatBoolean](../../api/Excel.modelformatboolean.md) object that represents formatting of type true/false in the data model. Read-only.|
|**[Model.ModelFormatCurrency Property (Excel)](../../api/Excel.model.modelformatcurrency.md)**|Returns a [ModelFormatCurrency](../../api/Excel.modelformatcurrency.md) object that represents formatting of type currency in the data model. Read-only.|
|**[Model.ModelFormatDate Property (Excel)](../../api/Excel.model.modelformatdate.md)**|Returns a [ModelFormatDate](../../api/Excel.modelformatdate.md) object that represents formatting of type date in the data model. Read-only.|
|**[Model.ModelFormatDecimalNumber Property (Excel)](../../api/Excel.model.modelformatdecimalnumber.md)**|Returns a [ModelFormatDecimalNumber](../../api/Excel.modelformatdecimalnumber.md) object that represents formatting of type decimal number in the data model. Read-only.|
|**[Model.ModelFormatGeneral Property (Excel)](../../api/Excel.model.modelformatgeneral.md)**|Returns a [ModelFormatGeneral](../../api/Excel.modelformatgeneral.md) object that represents formatting of type general in the data model. Read-only.|
|**[Model.ModelFormatPercentageNumber Property (Excel)](../../api/Excel.model.modelformatpercentagenumber.md)**|Returns a [ModelFormatPercentageNumber](../../api/Excel.modelformatpercentagenumber.md) object that represents formatting of type percentage number in the data model. Read-only.|
|**[Model.ModelFormatScientificNumber Property (Excel)](../../api/Excel.model.modelformatscientificnumber.md)**|Returns a [ModelFormatScientificNumber](../../api/Excel.modelformatscientificnumber.md) object that represents formatting of type scientific number in the data model. Read-only.|
|**[Model.ModelFormatWholeNumber Property (Excel)](../../api/Excel.model.modelformatwholenumber.md)**|Returns a [ModelFormatWholeNumber](../../api/Excel.modelformatwholenumber.md) object that represents formatting of type whole number in the data model. Read-only.|
|**[Model.ModelMeasures Property (Excel)](../../api/Excel.model.modelmeasures.md)**|Returns a [ModelMeasures](../../api/Excel.modelmeasures.md) object that represents the collection of model measures in the data model. Read-only.|
|**[ModelConnection.CalculatedMembers Property (Excel)](../../api/Excel.modelconnection.calculatedmembers.md)**|Returns a [CalculatedMembers](../../api/Excel.modelconnection.calculatedmembers.md) object that represents the calculated members in the model connection.|
|**[ModelFormatBoolean Object (Excel)](../../api/Excel.modelformatboolean.md)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatBoolean.Application Property (Excel)](../../api/Excel.modelformatboolean.application.md)**|When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an [Application](../../api/Excel.Application(object).md) object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatBoolean.Creator Property (Excel)](../../api/Excel.modelformatboolean.creator.md)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatBoolean.Parent Property (Excel)](../../api/Excel.modelformatboolean.parent.md)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatCurrency Object (Excel)](../../api/Excel.modelformatcurrency.md)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatCurrency.Application Property (Excel)](../../api/Excel.modelformatcurrency.application.md)**|When used without an object qualifier, this property returns an [Application](../../api/Excel.Application(object).md) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatCurrency.Creator Property (Excel)](../../api/Excel.modelformatcurrency.creator.md)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatCurrency.DecimalPlaces Property (Excel)](../../api/Excel.modelformatcurrency.decimalplaces.md)**|Specifies the number of decimal places after the dot. Read/write  **Long**.|
|**[ModelFormatCurrency.Parent Property (Excel)](../../api/Excel.modelformatcurrency.parent.md)**|Returns the parent object for the specified object. Read-only|
|**[ModelFormatCurrency.Symbol Property (Excel)](../../api/Excel.modelformatcurrency.symbol.md)**|Specifies the symbol to use to represent the currency. Read/write  **String**.|
|**[ModelFormatDate Object (Excel)](../../api/Excel.modelformatdate.md)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatDate.Application Property (Excel)](../../api/Excel.modelformatdate.application.md)**|When used without an object qualifier, this property returns an [Application](../../api/Excel.Application(object).md) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatDate.Creator Property (Excel)](../../api/Excel.modelformatdate.creator.md)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatDate.FormatString Property (Excel)](../../api/Excel.modelformatdate.formatstring.md)**|Specifies the date format, for example, " _dd/mm/yy_ ". Read/write **String**.|
|**[ModelFormatDate.Parent Property (Excel)](../../api/Excel.modelformatdate.parent.md)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatDecimalNumber Object (Excel)](../../api/Excel.modelformatdecimalnumber.md)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatDecimalNumber.Application Property (Excel)](../../api/Excel.modelformatdecimalnumber.application.md)**|When used without an object qualifier, this property returns an [Application](../../api/Excel.Application(object).md) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatDecimalNumber.Creator Property (Excel)](../../api/Excel.modelformatdecimalnumber.creator.md)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatDecimalNumber.DecimalPlaces Property (Excel)](../../api/Excel.modelformatdecimalnumber.decimalplaces.md)**|Specifies the number of decimal places after the dot. Read/write  **Long**.|
|**[ModelFormatDecimalNumber.Parent Property (Excel)](../../api/Excel.modelformatdecimalnumber.parent.md)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatDecimalNumber.UseThousandSeparator Property (Excel)](../../api/Excel.modelformatdecimalnumber.usethousandseparator.md)**|Specifies whether to display commas between thousands. Read/write  **Boolean**.|
|**[ModelFormatGeneral Object (Excel)](../../api/Excel.modelformatgeneral.md)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatGeneral.Application Property (Excel)](../../api/Excel.modelformatgeneral.application.md)**|When used without an object qualifier, this property returns an [Application](../../api/Excel.Application(object).md) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatGeneral.Creator Property (Excel)](../../api/Excel.modelformatgeneral.creator.md)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatGeneral.Parent Property (Excel)](../../api/Excel.modelformatgeneral.parent.md)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatPercentageNumber Object (Excel)](../../api/Excel.modelformatpercentagenumber.md)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatPercentageNumber.Application Property (Excel)](../../api/Excel.modelformatpercentagenumber.application.md)**|When used without an object qualifier, this property returns an [Application](../../api/Excel.Application(object).md) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatPercentageNumber.Creator Property (Excel)](../../api/Excel.modelformatpercentagenumber.creator.md)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatPercentageNumber.DecimalPlaces Property (Excel)](../../api/Excel.modelformatpercentagenumber.decimalplaces.md)**|Specifies the number of decimal places after the dot. Read/write  **Long**.|
|**[ModelFormatPercentageNumber.Parent Property (Excel)](../../api/Excel.modelformatpercentagenumber.parent.md)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatPercentageNumber.UseThousandSeparator Property (Excel)](../../api/Excel.modelformatpercentagenumber.usethousandseparator.md)**|Specifies whether to display commas between thousands. Read/write  **Boolean**.|
|**[ModelFormatScientificNumber Object (Excel)](../../api/Excel.modelformatscientificnumber.md)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatScientificNumber.Application Property (Excel)](../../api/Excel.modelformatscientificnumber.application.md)**|When used without an object qualifier, this property returns an [Application](../../api/Excel.Application(object).md) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatScientificNumber.Creator Property (Excel)](../../api/Excel.modelformatscientificnumber.creator.md)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatScientificNumber.DecimalPlaces Property (Excel)](../../api/Excel.modelformatscientificnumber.decimalplaces.md)**|Specifies the number of decimal places after the dot. Read/write  **Long**.|
|**[ModelFormatScientificNumber.Parent Property (Excel)](../../api/Excel.modelformatscientificnumber.parent.md)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatWholeNumber Object (Excel)](../../api/Excel.modelformatwholenumber.md)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatWholeNumber.Application Property (Excel)](../../api/Excel.modelformatwholenumber.application.md)**|When used without an object qualifier, this property returns an [Application](../../api/Excel.Application(object).md) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatWholeNumber.Creator Property (Excel)](../../api/Excel.modelformatwholenumber.creator.md)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatWholeNumber.Parent Property (Excel)](../../api/Excel.modelformatwholenumber.parent.md)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatWholeNumber.UseThousandSeparator Property (Excel)](../../api/Excel.modelformatwholenumber.usethousandseparator.md)**|Specifies whether to display commas between thousands. Read/write  **Boolean**.|
|**[ModelMeasure Object (Excel)](../../api/Excel.modelmeasure.md)**|Represents a single  **ModelMeasure** object in the [ModelMeasures](../../api/Excel.modelmeasures.md) collection.|
|**[ModelMeasure.Application Property (Excel)](../../api/Excel.modelmeasure.application.md)**|When used without an object qualifier, this property returns an [Application](../../api/Excel.Application(object).md) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelMeasure.AssociatedTable Property (Excel)](../../api/Excel.modelmeasure.associatedtable.md)**|Specifies the table that contains the model measure, as displayed in the  **Field List** task pane. Read/write[ModelTable](../../api/Excel.modeltable.md).|
|**[ModelMeasure.Creator Property (Excel)](../../api/Excel.modelmeasure.creator.md)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelMeasure.Delete Method (Excel)](../../api/Excel.modelmeasure.delete.md)**|Deletes the model measure from the data model.|
|**[ModelMeasure.Description Property (Excel)](../../api/Excel.modelmeasure.description.md)**|The description of the model measure. Read/write  **String**.|
|**[ModelMeasure.FormatInformation Property (Excel)](../../api/Excel.modelmeasure.formatinformation.md)**|The format of the model measure. Read/write  **Variant**.|
|**[ModelMeasure.Formula Property (Excel)](../../api/Excel.modelmeasure.formula.md)**|The Data Analysis Expressions (DAX) formula of the model measure. Read/write  **String**.|
|**[ModelMeasure.Name Property (Excel)](../../api/Excel.modelmeasure.name.md)**|The name of the model measure. Read/write  **String**.|
|**[ModelMeasure.Parent Property (Excel)](../../api/Excel.modelmeasure.parent.md)**|Returns the parent object for the specified object. Read-only.|
|**[ModelMeasures Object (Excel)](../../api/Excel.modelmeasures.md)**|Represents: a collection of  **ModelMeasure** objects.|
|**[ModelMeasures.Add Method (Excel)](../../api/Excel.modelmeasures.add.md)**|Adds a model measure to the model.|
|**[ModelMeasures.Application Property (Excel)](../../api/Excel.modelmeasures.application.md)**|When used without an object qualifier, this property returns an [Application](../../api/Excel.Application(object).md) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelMeasures.Count Property (Excel)](../../api/Excel.modelmeasures.count.md)**|Returns an integer that represents the number of objects in the collection.|
|**[ModelMeasures.Creator Property (Excel)](../../api/Excel.modelmeasures.creator.md)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelMeasures.Item Method (Excel)](../../api/Excel.modelmeasures.item.md)**|Returns a single object from a collection|
|**[ModelMeasures.Parent Property (Excel)](../../api/Excel.modelmeasures.parent.md)**|Returns the parent object for the specified object. Read-only.|
|**[ModelRelationships.DetectRelationships Method (Excel)](../../api/Excel.modelrelationships.detectrelationships.md)**|Detects model relationships in the specified [PivotTable](../../api/Excel.PivotTable.md).|
|**[PivotField.AutoGroup Method (Excel)](../../api/Excel.pivotfield.autogroup.md)**|Automatically groups the pivot fields in a pivot table.|
|**[Point.IsTotal Property (Excel)](../../api/Excel.point.istotal.md)**|**True** if the point represents a total. Read/write **Boolean**.|
|**[Queries Object (Excel)](../../api/Excel.queries.md)**|The collection of [WorkbookQuery](../../api/Excel.workbookquery.md) objects|
|**[Queries.Add Method (Excel)](../../api/Excel.queries.add.md)**|Adds a new [WorkbookQuery](../../api/Excel.workbookquery.md) object to the **Queries** collection.|
|**[Queries.Application Property (Excel)](../../api/Excel.queries.application.md)**|When used without an object qualifier, this property returns an [Application](../../api/Excel.Application(object).md) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[Queries.Count Property (Excel)](../../api/Excel.queries.count.md)**|Returns an integer that represents the number of objects in the collection.|
|**[Queries.Creator Property (Excel)](../../api/Excel.queries.creator.md)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[Queries.FastCombine Property (Excel)](../../api/Excel.queries.fastcombine.md)**|**True** to enable the fast combine feature, as long as the workbook is open. Read/write **Boolean**.|
|**[Queries.Item Method (Excel)](../../api/Excel.queries.item.md)**|Returns a single object from a collection.|
|**[Queries.Parent Property (Excel)](../../api/Excel.queries.parent.md)**|Returns the parent object for the specified object. Read-only.|
|**[Series.ParentDataLabelOption Property (Excel)](../../api/Excel.series.parentdatalabeloption.md)**|Specifies the parent data label option (banner, overlapping, or none) for the specified series within the chart group. Read/write [XLParentDataLabelOptions](../../api/Excel.xlparentdatalabeloptions.md).|
|**[Series.QuartileCalculationInclusiveMedian Property (Excel)](../../api/Excel.series.quartilecalculationinclusivemedian.md)**|**True** if the series uses an inclusive median quartile calculation method. Read/write **Boolean**.|
|**[SoundNote Object (Excel)](../../api/overview/Library-Reference.md)**|Represents a recorded sound note.|
|**[SoundNote.Application Property (Excel)](../../api/Excel.soundnote.application.md)**|When used without an object qualifier, this property returns an [Application](../../api/Excel.Application(object).md) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[SoundNote.Creator Property (Excel)](../../api/Excel.soundnote.creator.md)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[SoundNote.Parent Property (Excel)](../../api/Excel.soundnote.parent.md)**|Returns the parent object for the specified object. Read-only.|
|**[SoundNote.Delete Method (Excel)](../../api/Excel.soundnote.delete.md)**|Deletes the sound note.|
|**[SoundNote.Import Method (Excel)](../../api/Excel.soundnote.import.md)**|Imports the specified sound note.|
|**[SoundNote.Play Method (Excel)](../../api/Excel.soundnote.play.md)**|Plays the sound note.|
|**[SoundNote.Record Method (Excel)](../../api/Excel.soundnote.record.md)**|Records the sound note.|
|**[Workbook.CreateForecastSheet Method (Excel)](../../api/Excel.workbook.createforecastsheet.md)**|If you have historical time-based data, you can use  **CreateForecastSheet** to create a forecast. When you create a forecast, a new worksheet is created that contains a table of the historical and predicted values and a chart showing this. A forecast can help you predict things like future sales, inventory requirements, or consumer trends.|
|**[WorkbookQuery Object (Excel)](../../api/Excel.workbookquery.md)**|An object that represents a query that was created by Power Query.|
|**[WorkbookQuery.Application Property (Excel)](../../api/Excel.workbookquery.application.md)**|When used without an object qualifier, this property returns an [Application](../../api/Excel.Application(object).md) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[WorkbookQuery.Creator Property (Excel)](../../api/Excel.workbookquery.creator.md)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[WorkbookQuery.Delete Method (Excel)](../../api/Excel.workbookquery.delete.md)**|Deletes this query and its underlying connection and removes it from the  **Queries** collection.|
|**[WorkbookQuery.Description Property (Excel)](../../api/Excel.workbookquery.description.md)**|The description of the query. Read/write  **String**.|
|**[WorkbookQuery.Formula Property (Excel)](../../api/Excel.workbookquery.formula.md)**|The Power Query M formula for the object. Read-only  **String**.|
|**[WorkbookQuery.Name Property (Excel)](../../api/Excel.workbookquery.name.md)**|The name of the query. Read/write  **String**.|
|**[WorkbookQuery.Parent Property (Excel)](../../api/Excel.workbookquery.parent.md)**|Returns the parent object for the specified object. Read-only.|
|**[WorksheetFunction.Forecast_ETS Method (Excel)](../../api/Excel.worksheetfunction.forecast_ets.md)**|Calculates or predicts a future value based on existing (historical) values by using the AAA version of the Exponential Smoothing (ETS) algorithm. |
|**[WorksheetFunction.Forecast_ETS_ConfInt Method (Excel)](../../api/Excel.worksheetfunction.forecast_ets_confint.md)**|Returns a confidence interval for the forecast value at the specified target date.|
|**[WorksheetFunction.Forecast_ETS_Seasonality Method (Excel)](../../api/Excel.worksheetfunction.forecast_ets_seasonality.md)**|Returns the length of the repetitive pattern Excel detects for the specified time series.|
|**[WorksheetFunction.Forecast_ETS_STAT Method (Excel)](../../api/Excel.worksheetfunction.forecast_ets_stat.md)**|Returns a statistical value as a result of time series forecasting.|
|**[WorksheetFunction.Forecast_Linear Method (Excel)](../../api/Excel.worksheetfunction.forecast_linear.md)**|Calculates, or predicts, a future value by using existing values. The predicted value is a y-value for a given x-value. The known values are existing x-values and y-values, and the new value is predicted by using linear regression. You can use this function to predict future sales, inventory requirements, or consumer trends.|
|**[XlBinsType Enumeration (Excel)](../../api/Excel.xlbinstype.md)**|Constants passed to and returned by the [ChartGroup.BinsType](../../api/Excel.chartgroup.binstype.md) property.|
|**[XlForecastAggregation Enumeration (Excel)](../../api/Excel.xlforecastaggregation.md)**|Constants passed to various  **WorksheetFunction** and **Workbook** statistical forecasting methods.|
|**[XlForecastChartType Enumeration (Excel)](../../api/Excel.xlforecastcharttype.md)**|Constants passed to the [Workbook.CreateForecastSheet](../../api/Excel.workbook.createforecastsheet.md) Method.|
|**[XlForecastDataCompletion Enumeration (Excel)](../../api/Excel.xlforecastdatacompletion.md)**|Constants passed to various  **WorksheetFunction** and **Workbook** statistical forecasting methods.|
|**[XlParentDataLabelOptions Enumeration (Excel)](../../api/Excel.xlparentdatalabeloptions.md)**|Constants passed to and returned by the  **Series.ParentDataLabelOption** property.|

## Outlook





|**Name**|**Description**|
|:-----|:-----|
|**[ExchangeDistributionList.GetUnifiedGroup Method (Outlook)](../../api/Outlook.exchangedistributionlist.getunifiedgroup.md)**|Determines if the object is a unified group, by way of a call to [IsUnifiedGroup](../../api/Outlook.exchangedistributionlist.isunifiedgroup.md)|
|**[ExchangeDistributionList.GetUnifiedGroupFromStore Method (Outlook)](../../api/Outlook.exchangedistributionlist.getunifiedgroupfromstore.md)**|Determines if the object is a unified group (by way of a call to [IsUnifiedGroup](../../api/Outlook.exchangedistributionlist.isunifiedgroup.md)) and returns the  **Outlook.Folder** object associated with the group using the [GetUnifiedGroup](../../api/Outlook.exchangedistributionlist.getunifiedgroup.md) and **GetUnifiedGroupFromStore** methods.|
|**[ExchangeDistributionList.IsUnifiedGroup Method (Outlook)](../../api/Outlook.exchangedistributionlist.isunifiedgroup.md)**|Determines if the object is a unified group.|
|**[ExchangeUser.GetUnifiedGroup Method (Outlook)](../../api/Outlook.exchangeuser.getunifiedgroup.md)**|Determines if the object is a unified group, by way of a call to [IsUnifiedGroup](../../api/Outlook.exchangeuser.isunifiedgroup.md).|
|**[ExchangeUser.GetUnifiedGroupFromStore Method (Outlook)](../../api/Outlook.exchangeuser.getunifiedgroupfromstore.md)**|Determines if the object is a unified group, by way of a call to [IsUnifiedGroup](../../api/Outlook.exchangeuser.isunifiedgroup.md).|
|**[ExchangeUser.IsUnifiedGroup Method (Outlook)](../../api/Outlook.exchangeuser.isunifiedgroup.md)**|Determines if the object is a unified group.|
|**[Explorer.DisplayMode Property (Outlook)](../../api/Outlook.explorer.displaymode.md)**|Indicates the display mode: Normal, Portrait View, or Portrait Reading Pane.|
|**[Explorer.DisplayModeChange Event (Outlook)](../../api/Outlook.explorer.displaymodechange.md)**|Occurs when the user performs an action that changes the display mode. Possible modes include Normal, Portrait View, and Portrait Reading Pane.|
|**[Explorer.PreviewPane Property (Outlook)](../../api/Outlook.explorer.previewpane.md)**|The [PreviewPane](../../api/Outlook.previewpane.md) object displays content in a "single pane mode" by showing only the Preview Pane view.|
|**[ExplorerEvents_10.DisplayModeChange Method (Outlook)](../../api/Outlook.explorerevents_10.displaymodechange.md)**|Occurs when the user performs an action that changes the display mode. Possible modes include Normal, Portrait View, and Portrait Reading Pane.|
|**[OlDisplayMode Enumeration (Outlook)](../../api/Outlook.oldisplaymode.md)**|Describes the nature of the display mode. Possible modes include Normal, Portrait View, and Portrait Reading Pane.|
|**[OlUnifiedGroupFolderType Enumeration (Outlook)](../../api/Outlook.olunifiedgroupfoldertype.md)**|Specifies the folder to be obtained for unified groups. Because groups have both a mail and a calendar folder, you can specify either the  **olGroupMailFolder** or **olGroupCalendarFolder**.|
|**[OlUnifiedGroupType Enumeration (Outlook)](../../api/Outlook.olunifiedgrouptype.md)**|Specifies the group type as public or private for the [CreateUnifiedGroup](../../api/Outlook.store.createunifiedgroup.md) method.|
|**[PreviewPane Members (Outlook)](../../api/overview/Library-Reference.md)**|Displays content in a "single pane mode" by showing only the Preview Pane view.|
|**[PreviewPane Object (Outlook)](../../api/Outlook.previewpane.md)**|Displays content in a "single pane mode" by showing only the Preview Pane view.|
|**[PreviewPane.Application Property (Outlook)](../../api/Outlook.previewpane.application.md)**|Returns the [Application](../../api/Outlook.Application.md) object that represents the parent application (Outlook) for the [PreviewPane](../../api/Outlook.previewpane.md) Object. Read-only.|
|**[PreviewPane.Class Property (Outlook)](../../api/Outlook.previewpane.class.md)**|Returns a constant in the [OlObjectClass](../../api/Outlook.OlObjectClass.md) enumeration indicating the class of the [PreviewPane](../../api/Outlook.previewpane.md) Object. Read-only.|
|**[PreviewPane.Parent Property (Outlook)](../../api/Outlook.previewpane.parent.md)**|Returns the parent property for the [PreviewPane](../../api/Outlook.previewpane.md) Object. Read only.|
|**[PreviewPane.Session Property (Outlook)](../../api/Outlook.previewpane.session.md)**|Returns the [NameSpace](../../api/Outlook.NameSpace.md) for the current session. Read-only.|
|**[PreviewPane.WordEditor Property (Outlook)](../../api/Outlook.previewpane.wordeditor.md)**|Returns the Microsoft Word Document Object Model of the message being displayed. Read-only.|
|**[Store.CreateUnifiedGroup Method (Outlook)](../../api/Outlook.store.createunifiedgroup.md)**|Enables a unified group to be created.|
|**[Store.DeleteUnifiedGroup Method (Outlook)](../../api/Outlook.store.deleteunifiedgroup.md)**|Enables a unified group to be deleted.|

## Project





|**Name**|**Description**|
|:-----|:-----|
|**[Application.AddEngagement Method (Project)](../../api/Project.application.addengagement.md)**|Adds a  **Resource Plan** view, enabling users to display and edit engagement data to Project when connected to Project Online. Introduced in Office 2016.|
|**[Application.EngagementInfo Method (Project)](../../api/Project.application.engagementinfo.md)**|Displays the engagement information dialog box user interface for the  **Resource Plan** view. Introduced in Office 2016.|
|**[Application.GetDpiScaleFactor Method (Project)](../../api/Project.application.getdpiscalefactor.md)**|Indicates the  **DPI Scale Factor**, used for optimizing scale settings. Introduced in Office 2016.|
|**[Application.InsertTimelineBar Method (Project)](../../api/Project.application.addtimelinebar.md)**|Adds a  **timeline** bar to the view.|
|**[Application.Inspector Method (Project)](../../api/Project.application.inspector.md)**|Indicates the  **Task Inspector** for use with engagement data.|
|**[Application.LocaleName Method (Project)](../../api/Project.application.localename.md)**|Language name that is used by Project, such as en-us or za-ch.|
|**[Application.ProjectSummaryInfoEx Method (Project)](../../api/Project.application.projectsummaryinfoex.md)**|Returns information about project summary, including the Project Utilization type and Project Utilization date information.|
|**[Application.RefreshEngagementsForProject Method (Project)](../../api/Project.application.refreshengagementsforproject.md)**|Refreshes the engagements for the project using engagement state on the server.|
|**[Application.RemoveTimelineBar Method (Project)](../../api/Project.application.removetimelinebar.md)**|Removes a  **Timeline** bar from the view.|
|**[Application.SubmitAllEngagementsForProject Method (Project)](../../api/Project.application.submitallengagementsforproject.md)**|Submits all the engagements in the project to the resource manager for review.|
|**[Application.SubmitSelectedEngagementsForProject Method (Project)](../../api/Project.application.submitselectedengagementsforproject.md)**|Submits all selected engagements in the project to the resource manager for review.|
|**[Application.TaskOnTimelineEx Method (Project)](../../api/Project.application.taskontimelineex.md)**|Manages tasks on the Timeline pane or for a specified custom timeline, including specifying the bar that you want to add or remove.|
|**[Application.TimelineBarDateRange Method (Project)](../../api/Project.application.timelinebardaterange.md)**|Modifies the start and finish dates for a  **Timeline** bar.|
|**[Application.UpdateEngagementsForProject Method (Project)](../../api/Project.application.updateengagementsforproject.md)**|Update the Engagements for a Project.|
|**[Assignment.Compliant Property (Project)](../../api/Project.assignment.compliant.md)**|Gets the compliant for a task assignment in Project. Read-only.|
|**[Cell.Engagement Property (Project)](../../api/Project.cell.engagement.md)**|Gets or sets the engagement resource for a cell.|
|**[Chart Members (Project)](../../api/overview/Library-Reference.md)**|The  **Chart** object represents a chart on a report in Project.|
|**[Engagement Object (Project)](../../api/Project.engagement.md)**||
|**[Engagement.Application Property (Project)](../../api/Project.engagement.application.md)**||
|**[Engagement.Comments Property (Project)](../../api/Project.engagement.comments.md)**||
|**[Engagement.CommittedFinish Property (Project)](../../api/Project.engagement.committedfinish.md)**||
|**[Engagement.CommittedMaxUnits Property (Project)](../../api/Project.engagement.committedmaxunits.md)**||
|**[Engagement.CommittedStart Property (Project)](../../api/Project.engagement.committedstart.md)**||
|**[Engagement.CommittedWork Property (Project)](../../api/Project.engagement.committedwork.md)**||
|**[Engagement.CreatedDate Property (Project)](../../api/Project.engagement.createddate.md)**||
|**[Engagement.Delete Method (Project)](../../api/Project.engagement.delete.md)**||
|**[Engagement.DraftFinish Property (Project)](../../api/Project.engagement.draftfinish.md)**||
|**[Engagement.DraftMaxUnits Property (Project)](../../api/Project.engagement.draftmaxunits.md)**||
|**[Engagement.DraftStart Property (Project)](../../api/Project.engagement.draftstart.md)**||
|**[Engagement.DraftWork Property (Project)](../../api/Project.engagement.draftwork.md)**||
|**[Engagement.GetField Method (Project)](../../api/Project.engagement.getfield.md)**||
|**[Engagement.Guid Property (Project)](../../api/Project.engagement.guid.md)**||
|**[Engagement.Index Property (Project)](../../api/Project.engagement.index.md)**||
|**[Engagement.ModifiedByGuid Property (Project)](../../api/Project.engagement.modifiedbyguid.md)**||
|**[Engagement.ModifiedByName Property (Project)](../../api/Project.engagement.modifiedbyname.md)**||
|**[Engagement.ModifiedDate Property (Project)](../../api/Project.engagement.modifieddate.md)**||
|**[Engagement.Name Property (Project)](../../api/Project.engagement.name.md)**||
|**[Engagement.Parent Property (Project)](../../api/Project.engagement.parent.md)**||
|**[Engagement.ProjectGuid Property (Project)](../../api/Project.engagement.projectguid.md)**||
|**[Engagement.ProjectName Property (Project)](../../api/Project.engagement.projectname.md)**||
|**[Engagement.ProposedFinish Property (Project)](../../api/Project.engagement.proposedfinish.md)**||
|**[Engagement.ProposedMaxUnits Property (Project)](../../api/Project.engagement.proposedmaxunits.md)**||
|**[Engagement.ProposedStart Property (Project)](../../api/Project.engagement.proposedstart.md)**||
|**[Engagement.ProposedWork Property (Project)](../../api/Project.engagement.proposedwork.md)**||
|**[Engagement.ResourceGuid Property (Project)](../../api/Project.engagement.resourceguid.md)**||
|**[Engagement.ResourceID Property (Project)](../../api/Project.engagement.resourceid.md)**||
|**[Engagement.ResourceName Property (Project)](../../api/Project.engagement.resourcename.md)**||
|**[Engagement.ReviewedByGuid Property (Project)](../../api/Project.engagement.reviewedbyguid.md)**||
|**[Engagement.ReviewedByName Property (Project)](../../api/Project.engagement.reviewedbyname.md)**||
|**[Engagement.ReviewedDate Property (Project)](../../api/Project.engagement.revieweddate.md)**||
|**[Engagement.SetField Method (Project)](../../api/Project.engagement.setfield.md)**||
|**[Engagement.Status Property (Project)](../../api/Project.engagement.status.md)**||
|**[Engagement.SubmittedByGuid Property (Project)](../../api/Project.engagement.submittedbyguid.md)**||
|**[Engagement.SubmittedByName Property (Project)](../../api/Project.engagement.submittedbyname.md)**||
|**[Engagement.SubmittedDate Property (Project)](../../api/Project.engagement.submitteddate.md)**||
|**[EngagementComment Members (Project)](../../api/overview/Library-Reference.md)**||
|**[EngagementComment Object (Project)](../../api/Project.engagementcomment.md)**||
|**[EngagementComment Properties (Project)](../../api/overview/Library-Reference.md)**||
|**[EngagementComment.Application Property (Project)](../../api/Project.engagementcomment.application.md)**||
|**[EngagementComment.AuthorResEmail Property (Project)](../../api/Project.engagementcomment.authorresemail.md)**||
|**[EngagementComment.AuthorResGuid Property (Project)](../../api/Project.engagementcomment.authorresguid.md)**||
|**[EngagementComment.AuthorResName Property (Project)](../../api/Project.engagementcomment.authorresname.md)**||
|**[EngagementComment.CreatedDate Property (Project)](../../api/Project.engagementcomment.createddate.md)**||
|**[EngagementComment.Guid Property (Project)](../../api/Project.engagementcomment.guid.md)**||
|**[EngagementComment.Message Property (Project)](../../api/Project.engagementcomment.message.md)**||
|**[EngagementComment.Parent Property (Project)](../../api/Project.engagementcomment.parent.md)**||
|**[EngagementComments Members (Project)](../../api/overview/Library-Reference.md)**||
|**[EngagementComments Methods (Project)](../../api/overview/Library-Reference.md)**||
|**[EngagementComments Object (Project)](../../api/Project.engagementcomments.md)**||
|**[EngagementComments Properties (Project)](../../api/overview/Library-Reference.md)**||
|**[EngagementComments.Add Method (Project)](../../api/Project.engagementcomments.add.md)**||
|**[EngagementComments.Application Property (Project)](../../api/Project.engagementcomments.application.md)**||
|**[EngagementComments.Count Property (Project)](../../api/Project.engagementcomments.count.md)**||
|**[EngagementComments.Item Property (Project)](../../api/Project.engagementcomments.item.md)**||
|**[EngagementComments.Parent Property (Project)](../../api/Project.engagementcomments.parent.md)**||
|**[Engagements Members (Project)](../../api/overview/Library-Reference.md)**||
|**[Engagements Methods (Project)](../../api/overview/Library-Reference.md)**||
|**[Engagements Object (Project)](../../api/Project.engagements.md)**||
|**[Engagements Properties (Project)](../../api/overview/Library-Reference.md)**||
|**[Engagements.Add Method (Project)](../../api/Project.engagements.add.md)**||
|**[Engagements.Application Property (Project)](../../api/Project.engagements.application.md)**||
|**[Engagements.Count Property (Project)](../../api/Project.engagements.count.md)**||
|**[Engagements.Item Property (Project)](../../api/Project.engagements.item.md)**||
|**[Engagements.Parent Property (Project)](../../api/Project.engagements.parent.md)**||
|**[Engagements.UniqueID Property (Project)](../../api/Project.engagements.uniqueid.md)**||
|**[PjAssignmentWarnings Enumeration (Project)](../../api/Project.pjassignmentwarnings.md)**|Defines the different types of warnings that may appear on assignments triggering indicators in the indicator column in sheet views.|
|**[PjEngagementViolationType Enumeration (Project)](../../api/Project.pjengagementviolationtype.md)**|Defines the different types of engagement violation types triggering indicators in the indicator column in sheet views on tasks/resources and assignments. Used internally for setting the right violation types on tasks and resources.|
|**[PjEngagementWarnings Enumeration (Project)](../../api/Project.pjengagementwarnings.md)**|Defines the different types of warnings that may appear on engagements triggering indicators in the indicator column in sheet views.|
|**[PjResourceWarnings Enumeration (Project)](../../api/Project.pjresourcewarnings.md)**|Defines the different types of warnings that may appear on resources triggering indicators in the indicator column in sheet views. |
|**[Project.Engagements Property (Project)](../../api/Project.project.engagements.md)**|Returns the root object for all Engagement properties.|
|**[Project.LastWssSyncDate Property (Project)](../../api/Project.project.lastwsssyncdate.md)**|Returns the last date on which Project was synced with Wss. Read-only  **DateType**.|
|**[Project.Timeline Property (Project)](../../api/Project.project.timeline.md)**|Returns the root object for all Timeline properties. Read/write  **object**.|
|**[Project.UtilizationDate Property (Project)](../../api/Project.project.utilizationdate.md)**|Used for portfolio analysis, such as Project Plan, Resource Engagements, or Project Plan until. Read-only. Project Plan uses the project plan to calculate resource availability, Resource Engagements uses Resource Engagements, and Project Plan until is a combination of Project Plan and Resource Engagements.|
|**[Project.UtilizationType Property (Project)](../../api/Project.project.utilizationtype.md)**|If the Project.UtilizationType Property (Project) property is Project Plan until, this date is used to switch between using the project plan to calculate resource availability or when resource engagements are used. Read-only.|
|**[Resource.Compliant Property (Project)](../../api/Project.resource.compliant.md)**|**True** if the resource is compliant with its engagement. Read/write **Boolean**.|
|**[Resource.EngagementCommittedFinish Property (Project)](../../api/Project.resource.engagementcommittedfinish.md)**|Returns the committed finish date for the engagement. Read-only  **DateType**.|
|**[Resource.EngagementCommittedMaxUnits Property (Project)](../../api/Project.resource.engagementcommittedmaxunits.md)**|Returns the committed max units for the engagement. Read-only  **Integer**.|
|**[Resource.EngagementCommittedStart Property (Project)](../../api/Project.resource.engagementcommittedstart.md)**|Returns the committed start date for the engagement. Read-only  **DateType**.|
|**[Resource.EngagementCommittedWork Property (Project)](../../api/Project.resource.engagementcommittedwork.md)**|Returns the committed work for the engagement. Read-only  **Double**.|
|**[Resource.EngagementDraftFinish Property (Project)](../../api/Project.resource.engagementdraftfinish.md)**|Returns the draft finish date for the engagement. Read-only  **DateType**.|
|**[Resource.EngagementDraftMaxUnits Property (Project)](../../api/Project.resource.engagementdraftmaxunits.md)**|Returns the draft max units for the engagement. Read-only  **Integer**.|
|**[Resource.EngagementDraftStart Property (Project)](../../api/Project.resource.engagementdraftstart.md)**|Returns the draft start date for the engagement. Read-only  **DateType**.|
|**[Resource.EngagementDraftWork Property (Project)](../../api/Project.resource.engagementdraftwork.md)**|Returns the draft work for the engagement. Read-only  **Double**.|
|**[Resource.EngagementProposedFinish Property (Project)](../../api/Project.resource.engagementproposedfinish.md)**|Returns the proposed finish date for the engagement. Read-only  **DateType**.|
|**[Resource.EngagementProposedMaxUnits Property (Project)](../../api/Project.resource.engagementproposedmaxunits.md)**|Returns the proposed maximum units for the engagement. Read-only  **Integer**.|
|**[Resource.EngagementProposedStart Property (Project)](../../api/Project.resource.engagementproposedstart.md)**|Returns the proposed start date for the engagement. Read-only  **DateType**.|
|**[Resource.EngagementProposedWork Property (Project)](../../api/Project.resource.engagementproposedwork.md)**|Returns the proposed work for the engagement. Read-only  **Double**.|
|**[Resource.IsLocked Property (Project)](../../api/Project.resource.islocked.md)**|Indicates whether the resource is or is not locked. If resource is locked, you need an engagement for a resource. Read-only Return  **Boolean**.|
|**[Task.Compliant Property (Project)](../../api/Project.task.compliant.md)**||
|**[Timeline Members (Project)](../../api/overview/Library-Reference.md)**||
|**[Timeline Object (Project)](../../api/Project.timeline.md)**||
|**[Timeline Properties (Project)](../../api/overview/Library-Reference.md)**||
|**[Timeline.Application Property (Project)](../../api/Project.timeline.application.md)**|Gets the Project  **Application** object.|
|**[Timeline.BarCount Property (Project)](../../api/Project.timeline.barcount.md)**|Indicates the number of bars in a  **Timeline** view.|
|**[Timeline.FinishDate Property (Project)](../../api/Project.timeline.finishdate.md)**|Indicates the finish date for a  **Timeline** bar based on the input argument.|
|**[Timeline.Label Property (Project)](../../api/Project.timeline.label.md)**|Returns the timeline for the  **Timeline** object.|
|**[Timeline.StartDate Property (Project)](../../api/Project.timeline.startdate.md)**|Indicates the start date for a  **Timeline** bar based on the input argument.|

## PowerPoint





|**Name**|**Description**|
|:-----|:-----|
|**[ChartGroup.BinsCountValue Property (PowerPoint)](../../api/PowerPoint.chartgroup.binscountvalue.md)**|Specifies the number of bins in the histogram chart. Read/write  **Long**.|
|**[ChartGroup.BinsOverflowEnabled Property (PowerPoint)](../../api/PowerPoint.chartgroup.binsoverflowenabled.md)**|Specifies whether a bin for values above the ChartGroup.BinsOverflowValue Property (PowerPoint) is enabled. Read/write  **Boolean**.|
|**[ChartGroup.BinsOverflowValue Property (PowerPoint)](../../api/PowerPoint.chartgroup.binsoverflowvalue.md)**|If an [ChartGroup.BinsOverflowEnabled](../../api/PowerPoint.chartgroup.binsoverflowenabled.md) Property (PowerPoint) is **True**, specifies the value above which an overflow bin is displayed. Read/write **Double**.|
|**[ChartGroup.BinsType Property (PowerPoint)](../../api/PowerPoint.chartgroup.binstype.md)**|Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Read/write [XlBinsType](../../api/PowerPoint.xlbinstype.md) Enumeration (PowerPoint).|
|**[ChartGroup.BinsUnderflowEnabled Property (PowerPoint)](../../api/PowerPoint.chartgroup.binsunderflowenabled.md)**|Specifies whether a bin for values below the [ChartGroup.BinsUnderflowValue](../../api/PowerPoint.chartgroup.binsunderflowvalue.md) Property (PowerPoint) is enabled. Read/write **Boolean**.|
|**[ChartGroup.BinsUnderflowValue Property (PowerPoint)](../../api/PowerPoint.chartgroup.binsunderflowvalue.md)**|If [ChartGroup.BinsUnderflowEnabled](../../api/PowerPoint.chartgroup.binsunderflowenabled.md) Property (PowerPoint) is True, specifies the value below which an underflow bin is displayed. Read/write Double.|
|**[ChartGroup.BinWidthValue Property (PowerPoint)](../../api/PowerPoint.chartgroup.binwidthvalue.md)**|Specifies the number of points in each range. Read/write  **Double**.|
|**[DocumentWindow.ShowInsertAppDialog Method (PowerPoint)](../../api/PowerPoint.documentwindow.showinsertappdialog.md)**||
|**[Point.IsTotal Property (PowerPoint)](../../api/PowerPoint.point.istotal.md)**|**True** if the point represents a total. Read/write **Boolean**.|
|**[Series.ParentDataLabelOption Property (PowerPoint)](../../api/PowerPoint.series.parentdatalabeloption.md)**|Specifies the parent data label option (banner, overlapping, or none) for the specified series within the chart group. Read/write [XlParentDataLabelOptions](../../api/PowerPoint.xlparentdatalabeloptions.md) Enumeration (PowerPoint).|
|**[Series.QuartileCalculationInclusiveMedian Property (PowerPoint)](../../api/PowerPoint.series.quartilecalculationinclusivemedian.md)**|**True** if the series uses an inclusive median quartile calculation method. Read/write **Boolean**.|
|**[Shape.HasInkXML Property (PowerPoint)](../../api/PowerPoint.shape.hasinkxml.md)**|Returns an [MsoTriState](../../api/Office.MsoTriState.md) enumeration value that indicates whether the specified shape contains ink XML that can be retrieved via the [Shape.InkXML](../../api/PowerPoint.shape.inkxml.md) property. Read-only. An error is returned if the shape does not contain any ink XML.|
|**[Shape.InkXML Property (PowerPoint)](../../api/PowerPoint.shape.inkxml.md)**|Returns a  **String** that contains the InkActionML associated with the specified shape. Read-only. If the specified shape does not contain a ink object more than one ink object occurs , an error is returned.|
|**[Shape.IsNarration Property (PowerPoint)](../../api/PowerPoint.shape.isnarration.md)**|Specifies whether the specified shape range contains a narration. Read/write.|
|**[ShapeRange.HasInkXML Property (PowerPoint)](../../api/PowerPoint.shaperange.hasinkxml.md)**|Returns an [MsoTriState](../../api/Office.MsoTriState.md) enumeration value that indicates whether the specified shape range contains ink XML that can be retrieved via the [ShapeRange.InkXML](../../api/PowerPoint.shaperange.inkxml.md) property. Read-only. An error is returned if the shape range does not contain any ink XML.|
|**[ShapeRange.InkXML Property (PowerPoint)](../../api/PowerPoint.shaperange.inkxml.md)**|Returns a  **String** that contains the InkActionML associated with the specified shape range. Read-only. If the specified shape range does not contain a ink object more than one ink object occurs , an error is returned.|
|**[ShapeRange.IsNarration Property (PowerPoint)](../../api/PowerPoint.shaperange.isnarration.md)**|Specifies whether the specified shape range contains a narration. Read/write. |
|**[Shapes.AddInkShapeFromXML Method (PowerPoint)](../../api/PowerPoint.shapes.addinkshapefromxml.md)**|Creates an ink shape. Returns a [Shape](../../api/PowerPoint.Shape.md) object that represents the new ink shape.|
|**[SlideShowView.LaserPointerEnabled Property (PowerPoint)](../../api/PowerPoint.slideshowview.laserpointerenabled.md)**|Returns  **true** if the current slide show pointer is a laser pointer. This property is applicable only while the slide show is running. Read/write. This property allows a user to programmatically query and set the state of the pointer shown during slide show. The property will return false for all other pointer types. Users can also change the state of the current pointer by setting this property to **true** to turn on the laser pointer or **false** to turn off the laser pointer.|
|**[XlBinsType Enumeration (PowerPoint)](../../api/PowerPoint.xlbinstype.md)**|Constants passed to and returned by the [ChartGroup.BinsType](../../api/Excel.chartgroup.binstype.md) property.|
|**[XlParentDataLabelOptions Enumeration (PowerPoint)](../../api/PowerPoint.xlparentdatalabeloptions.md)**|Constants passed to and returned by the  **Series.ParentDataLabelOption** property.|

## Visio





|**Name**|**Description**|
|:-----|:-----|
|**[Document.Permission Property (Visio)](../../api/Visio.document.permission.md)**||
|**[IVInvisibleApp.Application Property (Visio)](../../api/overview/Library-Reference.md)**||
|**[IVKeyboardEvent.Application Property (Visio)](../../api/overview/Library-Reference.md)**||
|**[IVMouseEvent.Application Property (Visio)](../../api/overview/Library-Reference.md)**||
|**[Master.VisualBoundingBox Method (Visio)](../../api/Visio.master.visualboundingbox.md)**|Returns the bounding rectangle of the virtual container that has all the shapes of the given master.|
|**[Page.VisualBoundingBox Method (Visio)](../../api/Visio.page.visualboundingbox.md)**|Returns the bounding rectangle of the virtual container that has all the shapes of the given page.|
|**[Selection.VisualBoundingBox Method (Visio)](../../api/Visio.selection.visualboundingbox.md)**|Returns the bounding rectangle of the virtual container that has all the shapes of the given selection.|
|**[Shape.VisualBoundingBox Method (Visio)](../../api/Visio.shape.visualboundingbox.md)**|Returns the bounding rectangle of the given shape.|
|**[ValidationIssues.Stat Property (Visio)](../../api/Visio.validationissues.stat.md)**||
|**[VisColoringMethod Enumeration (Visio)](../../api/overview/Library-Reference.md)**||
|**[VisRecordsetFieldStatus Enumeration (Visio)](../../api/overview/Library-Reference.md)**||

## Word





|**Name**|**Description**|
|:-----|:-----|
|**[ChartGroup.BinsCountValue Property (Word)](../../api/Word.chartgroup.binscountvalue.md)**|Specifies the number of bins in the histogram chart. Read/write  **Long**.|
|**[ChartGroup.BinsOverflowEnabled Property (Word)](../../api/Word.chartgroup.binsoverflowenabled.md)**|Specifies whether a bin for values above the [BinsOverflowValue](../../api/Excel.chartgroup.binsoverflowvalue.md) is enabled. Read/write **Boolean**.|
|**[ChartGroup.BinsOverflowValue Property (Word)](../../api/Word.chartgroup.binsoverflowvalue.md)**|If an [BinsOverflowEnabled](../../api/Excel.chartgroup.binsoverflowenabled.md) is **True**, specifies the value above which an overflow bin is displayed. Read/write **Double**.|
|**[ChartGroup.BinsType Property (Word)](../../api/Word.chartgroup.binstype.md)**|Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Read/write [XlBinsType](../../api/Word.xlbinstype.md).|
|**[ChartGroup.BinsUnderflowEnabled Property (Word)](../../api/Word.chartgroup.binsunderflowenabled.md)**|Specifies whether a bin for values below the [BinsUnderflowValue](../../api/Word.chartgroup.binsunderflowvalue.md) is enabled. Read/write **Boolean**.|
|**[ChartGroup.BinsUnderflowValue Property (Word)](../../api/Word.chartgroup.binsunderflowvalue.md)**|If an [BinsUnderflowEnabled](../../api/Word.chartgroup.binsunderflowenabled.md) is **True**, specifies the value below which an underflow bin is displayed. Read/write **Double**.|
|**[ChartGroup.BinWidthValue Property (Word)](../../api/Word.chartgroup.binwidthvalue.md)**|Specifies the number of points in each range. Read/write  **Double**.|
|**[CoAuthUpdates Object (Word)](../../api/Word.coauthupdates.md)**|A collection of [CoAuthUpdate](../../api/Word.CoAuthUpdate.md) objects that represent the updates that were merged into the document at the last explicit save.|
|**[Options.UseLocalUserInfo Property (Word)](../../api/Word.options.uselocaluserinfo.md)**||
|**[Point.IsTotal Property (Word)](../../api/Word.point.istotal.md)**|**True** if the point represents a total. Read/write **Boolean**.|
|**[Series.ParentDataLabelOption Property (Word)](../../api/Word.series.parentdatalabeloption.md)**|Specifies the parent data label option (banner, overlapping, or none) for the specified series within the chart group. Read/write [XLParentDataLabelOptions](../../api/Word.xlparentdatalabeloptions.md).|
|**[Series.QuartileCalculationInclusiveMedian Property (Word)](../../api/Word.series.quartilecalculationinclusivemedian.md)**|**True** if the series uses an inclusive median quartile calculation method. Read/write **Boolean**.|
|**[XlBinsType Enumeration (Word)](../../api/Word.xlbinstype.md)**|Constants passed to and returned by the [ChartGroup.BinsType](../../api/Word.chartgroup.binstype.md) property.|
|**[XlParentDataLabelOptions Enumeration (Word)](../../api/Word.xlparentdatalabeloptions.md)**|Constants passed to and returned by the  **Series.ParentDataLabelOption** property.|

