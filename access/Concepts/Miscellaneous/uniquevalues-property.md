---
title: UniqueValues property
ROBOTS: INDEX
keywords: vbaac10.chm4531
f1_keywords:
- vbaac10.chm4531
ms.prod: access
api_name:
- Access.UniqueValues
ms.assetid: 41e403b9-a94a-252c-7561-51ac2df4cf0c
ms.date: 06/08/2017
localization_priority: Normal
---


# UniqueValues property

**Applies to:** Access 2013 | Access 2016

You can use the **UniqueValues** property when you want to omit records that contain duplicate data in the fields displayed in Datasheet view. For example, if a query's output includes more than one field, the combination of values from all fields must be unique for a given record to be included in the results.

> [!NOTE] 
> The **UniqueValues** property applies only to append and make-table action queries and select queries.


## Setting

The **UniqueValues** property uses the following settings.

|Setting|Description|
|:-----|:-----|
|Yes|Displays only the records in which the values of all fields displayed in Datasheet view are unique.|
|No|(Default) Displays all records.|

You can set the **UniqueValues** property in the query's property sheet or in SQL view of the Query window.

> [!NOTE] 
> You can set this property when you create a new query by using an SQL statement. The DISTINCT predicate corresponds to the **UniqueValues** property setting. The DISTINCTROW predicate corresponds to the **UniqueValues** property setting.


## Remarks

When you set the **UniqueValues** property to Yes, the results of the query aren't updatable and won't reflect subsequent changes made by other users.

The **UniqueValues** and **UniqueRecords** properties are related in that only one of them can be set to Yes at a time. When you set the **UniqueValues** property to Yes, for example, Microsoft Access automatically sets the **UniqueRecords** property to No. You can, however, set both of them to No. When both properties are set to No, all records are returned.

If you want to count the number of instances of a value in a field, create a totals query.


## Example

The SELECT statement in this example returns a list of the countries/regions in which there are customers. Because there may be many customers in each country/region, many records could have the same country/region in the Customers table. However, each country/region is represented only once in the query results.

This example uses the Customers table, which contains the following data.

|**Country/Region**|**Company name**|
|:-----|:-----|
|Brazil|Familia Arquibaldo|
|Brazil|Gourmet Lanchonetes|
|Brazil|Hanari Carnes|
|France|Du monde entier|
|France|Folies gourmandes|
|Germany|Frankenversand|
|Ireland|Hungry Owl All-Night Grocers|

This SQL statement returns the countries/regions in the following table:

```sql
SELECT DISTINCT Customers.Country 
FROM Customers; 

```

|**Countries/Regions returned**|
|:-----|
|Brazil|
|France|
|Germany|
|Ireland|

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]