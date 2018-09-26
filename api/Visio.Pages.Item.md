---
title: Pages.Item Property (Visio)
keywords: vis_sdr.chm11013765
f1_keywords:
- vis_sdr.chm11013765
ms.prod: visio
api_name:
- Visio.Pages.Item
ms.assetid: c52ace02-486f-d50b-caf5-109b78008d77
ms.date: 06/08/2017
---


# Pages.Item Property (Visio)

Returns an item from a collection. The  **Item** property is the default property for all collections. Read-only.


## Syntax

 _expression_. `Item`( `_NameUIDOrIndex_` )

 _expression_ A variable that represents a [Pages](./Visio.Pages.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NameUIDOrIndex_|Required| **Variant**|Contains the name, unique ID, or index of the object to retrieve.|

### Return value

Page


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statements are equivalent to the syntax example given above:


```vb
objRet = object(index)  
objRet = object(stringExpression) 

```

You can retrieve an object in an  **Addons** , **Documents** , **Fonts** , **Hyperlinks** , **Layers** , **Masters** , **MasterShortcuts** , **OLEObjects** , **Pages** , **Shapes** , or **Styles** collection by passing the object's name as a string expression in a **Variant** .

For more information about passing ID strings to the  **Item** property, see the topic for the **UniqueID** property in this Automation Reference.


 **Note**  

Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **Item** property to access an object in the **Masters** , **Pages** , **Shapes** , **Styles** , **Layers** , or **MasterShortcuts** collection by using its local name. Use the **ItemU** property to access an object from one of these collections by using the object's universal name.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVPages.this[object]**
    

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Item** property to get a **Page** object from the **Pages** collection of the active document, and all the **Shape** objects in the **Shapes** collection of the **Page** object. It prints the names of all shapes on Page1 in the Immediate window.

Before running this macro, make sure that the active document has shapes on Page1.




```vb
 
Public Sub Item_Example() 
  
    Dim intCounter As Integer 
    Dim intShapeCount As Integer 
    Dim vsoShapes As Visio.Shapes  
 
    Set vsoShapes = ActiveDocument.Pages.Item(1).Shapes  
 
    Debug.Print "Shape Name List For..." 
    Debug.Print "Document: "; ActiveDocument.Name  
    Debug.Print "Page: "; ActiveDocument.Pages.Item(1).Name  
 
    intShapeCount = vsoShapes.Count  
 
    If intShapeCount > 0 Then 
        For intCounter = 1 To intShapeCount  
            Debug.Print " "; vsoShapes.Item(intCounter).Name  
        Next intCounter  
    Else 
        Debug.Print " No Shapes On Page"  
    End If   
 
End Sub
```


