---
title: DoCmd.SearchForRecord method (Access)
keywords: vbaac10.chm5765
f1_keywords:
- vbaac10.chm5765
ms.prod: access
api_name:
- Access.DoCmd.SearchForRecord
ms.assetid: eb7a82b0-1ecb-cbfe-94b0-e2d6742de8b4
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.SearchForRecord method (Access)

You can use the **SearchForRecord** method to search for a specific record in a table, query, form, or report.


## Syntax

_expression_.**SearchForRecord** (_ObjectType_, _ObjectName_, _Record_, _WhereCondition_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Optional|**[AcDataObjectType](Access.AcDataObjectType.md)**|An **AcDataObjectType** constant that specifies the type of database object in which you are searching. The default value is **acActiveDataObject**.|
| _ObjectName_|Optional|**Variant**|The name of the database object that contains the record to search for.|
| _Record_|Optional|**[AcRecord](Access.AcRecord.md)**|An **AcRecord** constant that specifies the starting point and direction of the search. The default value is **acFirst**.|
| _WhereCondition_|Optional|**Variant**|A string used to locate the record. It is like the WHERE clause in an SQL statement, but without the word WHERE.|

## Remarks

In cases where more than one record matches the criteria in the _WhereCondition_ argument, the following factors determine which record is found:
    
- The _Record_ argument setting.
    
- The sort order of the records. For example, if the _Record_ argument is set to **acFirst**, changing the sort order of the records might change which record is found.
    
The object specified in the _ObjectName_ argument must be open before this action is run. Otherwise, an error occurs.
    
If the criteria in the _WhereCondition_ argument are not met, no error occurs and the focus remains on the current record.
    
When searching for the previous or next record, the search does not "wrap" when it reaches the end of the data. If there are no further records that match the criteria, no error occurs and the focus remains on the current record. To confirm that a match was found, you can enter a condition for the next action, and make the condition the same as the criteria in the _WhereCondition_ argument.
    
The **SearchForRecord** method is similar to the **[FindRecord](Access.DoCmd.FindRecord.md)** method, but **SearchForRecord** has more powerful search features. The **FindRecord** method is primarily used for finding strings, and it duplicates the functionality of the **Find** dialog box. The **SearchForRecord** method uses criteria that are more like those of a filter or an SQL query. 

The following list demonstrates some things that you can do with the **SearchForRecord** method:
    
- You can use complex criteria in the _WhereCondition_ argument, such as `Description = "Beverages" and CategoryID = 11`.
 
- You can refer to fields that are in the record source of a form or report but are not displayed on the form or report. In the preceding example, neither `Description` nor `CategoryID` must be displayed on the form or report for the criteria to work.
    
- You can use logical operators, such as **<**, **>**, **AND**, **OR**, and **BETWEEN**. The **FindRecord** method only matches strings that equal, start with, or contain the string being searched for.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
