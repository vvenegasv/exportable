# Requerimientos
**Tener una clase plana**   
`public class DummyPerson`   
`{`   
    `    public int Edad { get; set; }`   
    `    public string Nombre { get; set; }`   
`}`

**Agregar los siguientes using**:   
`using Infodinamica.Framework.Exportable.Engines;`  
`using Infodinamica.Framework.Exportable.Engines.Excel;`

# Exportación Simple   
**Utilizar la interfaz IExportEngine**  
`IList<DummyPerson> dummyPeople = new List<DummyPerson>();`    
`//Add data to dummyPeople...`   
`IExportEngine engine = new ExcelExportEngine();`  
`engine.AddData(dummyPeople); `  
`MemoryStream memory = engine.Export();`  


# Especificar versión de Excel
**Utilizar la interfaz IExcelExportEngine**  
`IList<DummyPerson> dummyPeople = new List<DummyPerson>();`    
`//Add data to dummyPeople...`   
`IExcelExportEngine engine = new ExcelExportEngine();`  
`engine.SetFormat(ExcelVersion.XLS);`    
`engine.AddData(dummyPeople); `  
`MemoryStream memory = engine.Export();` 


# Especificar nombres de columnas, orden de campos y formatos
**Primero, especificar el siguiente using**    
`using Infodinamica.Framework.Exportable.Attribute;`      

**Segundo, especificar los atributos con "Exportable"**    
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
