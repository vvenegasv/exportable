# Changes in version 2.0
* Version 2.0 has a shorter namespace, from `Infodinamica.Framework.Exportable` to `Exportable`
* Version 2.0 don't requieres set the `Key` parameter in `AddContainer` method
* Version 2.0 support ignore and rename columns on runtime
* Version 2.0 don't use Infodinamica.Framework.Core
* Version 2.0 use the latest stable NPOI release (2.3.0)

# Installation with NUGET
`Install-Package Infodinamica.Framework.Exportable`

# Requirements
**Have a plain class**   
`public class DummyPerson`   
`{`   
    `    public int Edad { get; set; }`   
    `    public string Nombre { get; set; }`   
`}`

**Add this usings**:   
`using Exportable.Engines;`  
`using Exportable.Engines.Excel;`

# Simple Export   
**Use the IExportEngine interface**  
`IList<DummyPerson> dummyPeople = new List<DummyPerson>();`    
`//Add data to dummyPeople...`   
`IExportEngine engine = new ExcelExportEngine();`  
`engine.AddData(dummyPeople); `  
`MemoryStream memory = engine.Export();`  


# Set Excel version
**Use the IExcelExportEngine interface**  
`IList<DummyPerson> dummyPeople = new List<DummyPerson>();`    
`//Add data to dummyPeople...`   
`IExcelExportEngine engine = new ExcelExportEngine();`  
`engine.SetFormat(ExcelVersion.XLS);`    
`engine.AddData(dummyPeople); `  
`MemoryStream memory = engine.Export();` 


# Set columns name's, order and format
**First, add this using**    
`using Exportable.Attribute;`      

**Second, set "Exportable" attributes**    
`public class DummyPerson`   
`{`   
`[Exportable(3, "Full Name", FieldValueType.Text)]`    
`public string Name { get; set; }`    

`[Exportable(1, "Birth Date", FieldValueType.Date, "MM-yyyy")]`   
`public DateTime BirthDate { get; set; }`   

`[Exportable(2, "How Many Years", FieldValueType.Numeric, "#0")]`   
`public int Age { get; set; }`   

`[Exportable(4, "Is Adult", FieldValueType.Bool)]`    
`public bool IsAdult { get; set; }  `   
`}`
