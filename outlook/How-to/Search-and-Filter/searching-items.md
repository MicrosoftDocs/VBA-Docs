---
title: Searching Items
ms.prod: outlook
ms.assetid: f0c24b9d-160e-3218-6979-2071a3135bfc
ms.date: 06/08/2017
localization_priority: Normal
---


# Searching Items

This topic describes the entry points to search items in folders and access to search results.

||[Application.AdvancedSearch](../../../api/Outlook.Application.AdvancedSearch.md)|[Explorer.Search](../../../api/Outlook.Explorer.Search.md)|[Search.GetTable](../../../api/Outlook.Search.GetTable.md)|
|:-----|:-----|:-----|:-----|
|**Purpose**|Provides programmatic search on items in a specified folder based on a filter.|Performs a programmatic content indexer search that is analogous to a user executing a search from the Outlook user interface.|Provides an efficient way to access items (in a **[Table](../../../api/Outlook.Table.md)**) returned by a prior **Application.AdvancedSearch**. This entry point does not carry out a separate search.|
|**Scope of search**|Folder specified as a search parameter.|Determined by the parameter _SearchAllItems_. If _SearchAllItems_ is True, the method will search across all folders that have the same folder type as the current folder (specified by the **[DefaultItemType](../../../api/Outlook.Folder.DefaultItemType.md)** property of **[Explorer.CurrentFolder](../../../api/Outlook.Explorer.CurrentFolder.md)**) and all stores that have been selected for search in the Search Options dialog box. <br/><br/>If _SearchAllItems_ is False, the method will search only the folder represented by **Explorer.CurrentFolder**.|Because the **[Search](../../../api/Outlook.Search.md)** object is returned from a prior **Application.AdvancedSearch** operation, the scope of the search associated with **Search.GetTable** is that of the prior **Application.AdvancedSearch**.|
|**Search filter**|In DAV Searching and Locating (DASL) syntax.|Any valid keywords that are supported by Outlook search in the user interface. Search phrases are delimited by double quotes and can be concatenated to form a single search filter string.|Similar to the scope, the filter of the search associated with **Search.GetTable** is the filter parameter of the prior **Application.AdvancedSearch**.|
|**Search completion**|Use the **[AdvancedSearchComplete](../../../api/Outlook.Application.AdvancedSearchComplete.md)** event to determine when a given search has completed.|Does not provide a callback to indicate search completion.|Search is completed in the prior  **Application.AdvancedSearch**. **Search.GetTable** only returns the search results.|


## Search results

Access the search results by one of these means:

- **[Search.Results](../../../api/Outlook.Search.Results.md)** contains the search results as a **[Results](../../../api/Outlook.Results.md)** collection. Each item in the collection contains the full set of item properties.

- **Search**. **Save** saves the results to a search folder.

- **[Search.GetTable](../../../api/Outlook.Search.GetTable.md)** also returns the same set of items as in the **Results** collection, but each item  will contain only a default set of properties and therefore generally offers better performance.

Search results are displayed in Explorer for the current folder. To remove any search results in Explorer, call **[Explorer.ClearSearch](../../../api/Outlook.Explorer.ClearSearch.md)**. 

Search results are returned in a **Table** that includes the same set of items returned from the prior **Application.AdvancedSearch**. Because the **Table** only includes a limited set of properties per item, this is generally a more efficient way to access search results. To include properties other than the default in the search results, use **[Columns.Add](../../../api/Outlook.Columns.Add.md)** to get an updated **Table**. 

Because the item's Entry ID is one of the returned properties, you can also use **[GetItemFromID](../../../api/Outlook.NameSpace.GetItemFromID.md)** to obtain the item object, and access other item properties like **Body** and **AutoResolvedWinner** that are not supported by the **Table** object for performance reasons.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]