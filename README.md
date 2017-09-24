# 1. Changes in version 2.0
* Version 2.0 has a shorter namespace, from `Infodinamica.Framework.Exportable` to `Exportable`
* Version 2.0 don't requieres set the `Key` parameter in `AddContainer` method
* Version 2.0 support ignore and rename columns on runtime
* Version 2.0 don't use Infodinamica.Framework.Core
* Version 2.0 use the latest stable NPOI release (2.3.0)

# 2. Installation with NUGET
`Install-Package Infodinamica.Framework.Exportable`

# 3. Requirements
### 3.1 Have a plain class 
``` c#
public class DummyPerson
{
    public int Edad { get; set; }
    public string Nombre { get; set; }
}
```



# 4. Export

## 4.1. Namespaces
These are the commonly namespace required's to export:
``` c#
using Exportable.Engines;
using Exportable.Engines.Excel;
```

## 4.2. Simple Export  
``` c#
IList<DummyPerson> dummyPeople = new List<DummyPerson>();
//Add data to dummyPeople...
IExportEngine engine = new ExcelExportEngine();
engine.AddData(dummyPeople); 
MemoryStream memory = engine.Export();
```

## 4.3. Specifying the Excel version
``` c#
IList<DummyPerson> dummyPeople = new List<DummyPerson>();
//Add data to dummyPeople...
IExcelExportEngine engine = new ExcelExportEngine();
engine.SetFormat(ExcelVersion.XLS);
engine.AddData(dummyPeople); 
MemoryStream memory = engine.Export();
```

## 4.4. Set columns name's, order and format
### 4.4.1. First, add this using 
``` c#
using Exportable.Attribute;
```

### 4.4.2. Second, set "Exportable" attributes  
``` c#
public class DummyPerson
{   
    [Exportable(3, "Full Name", FieldValueType.Text)]  
    public string Name { get; set; }
    
    [Exportable(1, "Birth Date", FieldValueType.Date, "MM-yyyy")]   
    public DateTime BirthDate { get; set; } 
    
    [Exportable(2, "How Many Years", FieldValueType.Numeric, "#0")]   
    public int Age { get; set; }
    
    [Exportable(4, "Is Adult", FieldValueType.Bool)]    
    public bool IsAdult { get; set; }  
    
    [Exportable(IsIgnored = true)]
    public string ThisColumnWillBeIgnored { get; set; }
}
```

## 4.5. Override column names
You can override the column name with `AddColumnsNames` method. You need a plain class (in this case `DummyPersonWithAttributes`) and specify the new column name. Please note that this configuration overrides the column name specified in the class attribute
``` c#
IExportEngine engine = new ExcelExportEngine();
var key = engine.AddData(dummyPeople);
engine.AddColumnsNames<DummyPersonWithAttributes>(key, x => x.Name, "this is a new name LOL!");
var stream = engine.Export();
```

## 4.6. Ignore column names on runtime
You can ignore columns with `AddIgnoreColumns` method. You need a plain class (in this case `DummyPersonWithAttributes`) and specify the column that you want to ignore. Please note that this configuration overrides the column ignore specified in the class attribute
``` c#
IExportEngine engine = new ExcelExportEngine();
var key = engine.AddData(dummyPeople);
engine.AddIgnoreColumns<DummyPersonWithAttributes>(key, x => x.Name);
var stream = engine.Export();
```

## 4.7. Specify a column name by resource
You can handle internationalization column names, by using resources in attributes

### 4.7.1. Plain class
``` c#
class DummyPersonWithAttributesAndResource
{
    [Exportable(Position = 0, HeaderName = "Header1", ResourceType = typeof(res), TypeValue = FieldValueType.Text)]
    public string Name { get; set; }

    [Exportable(1, "Header2", FieldValueType.Date, "MM-yyyy", ResourceType = typeof(res))]
    public DateTime BirthDate { get; set; }

    [Exportable(IsIgnored = true)]
    public int Age { get; set; }

    [Exportable(3, "Is Adult", FieldValueType.Bool)]
    public bool IsAdult { get; set; }
}
```

### 4.7.2. Resource
Name     | Value
---------|------------
Header1  | Column 1
Header2  | Column 2

# Import

## Simple Import
