# ExcelReaders

## Tabla de Contenido

- [Motivación](#motivación)
- [Configuración](#configuración)
- [ExcelMappingSection](#excelmappingsection)
 - [ExcelMapping](#excelmapping)
 - [Mapper](#mapper)
 - [MapElement](#mapelemment)
 - [Sheet](#sheet)
- [Ejemplos de Configuración](#ejemplos-de-configuración)
- [Ejemplos de Uso](#ejemplos-de-uso)

## Motivación

Esta librería se creó con el objetivo de simplificar el accesos a datos en archivos de Excel.
La librería usa [NPOI](https://npoi.codeplex.com/) para la lectura de datos que entre otras cosas permite evitar las dependencias con 
librerias de Microsoft que tantos dolores de cabeza dan.

Es común importar datos desde Excel y utilizar clases para mapearlos,  existen diversas implementaciónes ya existentes
como [NPOI.Mapper](https://github.com/donnytian/Npoi.Mapper), que aunque son sencillas de usar, son pesadas y contienen múltiples fallos al momento de mapear los datos
a objetos de una clase en particular.

## Configuración

Para evitar la reflexión directa y darle velocidad al mapeo utilice una libreria de configuración, basta con agregar una sección
de configuración en el app.config o web.config para que los datos se importen en una ConcurrentBag<T> de acuerdo al los mappings configurados.

```
ProyectoWeb-Ejecutable/
	web.config ó app.config
	ExcelMappingSection.config
	
```

La sección puede configurarse de la siguiente manera:

```
<?xml version="1.0" encoding="utf-8" ?>
<configuration>
  <configSections>
    <section name="ExcelMappingSection" type="ExcelReaders.Configuration.ExcelMappingSection, ExcelReaders.Configuration, Version=1.0.0.0, Culture=neutral" />
  </configSections>

  <ExcelMappingSection configSource="ExcelMappingSection.config" />

```

## ExcelMappingSection


### ExcelMapping

Un ExcelMapping puede contener múltiples mappers y sheets. 

* **Name**: El nombre con que se identifica, este nombre permite asociar la configuración con el ExcelReader especifico.

### Mapper

Un mapper contiene MapElements que son asociaciones de los campos de la clase con las columnas de la hoja de Excel.

* **Name**: Es el nombre único que identifica el mapper este permitira asociarlo con una o varias sheets en la configuración.


### MapElement

* **NoColumn** : Es el número de columna de la hoja de Excel.	
* **Attribute**: Es el nombre del campo en la base de datos, funciona como identificador asi que no debe repetirse. 
* **Ignore**: Cuando esta establecido a true este campo no se mapea, es necesario establecer NoColumn a vacio.
* **Default**: El valor que se indica se establecerá sobre este campo independientemente de lo que se haya señalado para la hoja de Excel.
* **Format**: Este campo puede contener cualquier formato valido de C# sea prestalecido o personalizado, cuando se usa el tipo de campo a mapear deberá ser de tipo **string**.

### Sheet

* **NoSheet**: Número de hoja de excel, funciona como identificador así que no debe repetirse.
* **SheetName**: El nombre de la hoja en el archivo de Excel.
* **RowNumberStartData**: Fila donde empiezan los datos.
* **RowNumberStopData**: Fila donde terminan los datos.
* **Map**: Nombre del Mappper relacionado.


## Ejemplos de Configuración

En el código de la clase a mapear es el siguiente:

**BaseResult.cs**
```
public class BaseResult
{
	//Este campo se llena por medio de CalculateFields del ExcelReader
	//es usado en la búsqueda.
	public string Key { get; set; }
        
	public DateTime? FechaInicioVigencia { get; set; }

	public DateTime? FechaFinVigencia { get; set; }

	public string Version { get; set; }       
}
```


**AduanaResult.cs**

```
public class AduanaResult : BaseResult
	{
        public string Clave { get; set; }
        
        public override string ToString()
        {
            return $"{Clave}-{FechaInicioVigencia}-{Version}";
        }
    }
}

```

La configuración para mapear de la hoja de Excel a la clases es como sigue:

```
<ExcelMappingSection>
  <ExcelMappings>
    <ExcelMapping Name="Aduana">
      <Mappers>
        <Mapper Name="MapAduana">
          <MapElements>
            <MapElement NoColumn="0" Attribute="Clave" Ignore="false" Default="" Format="00"/>            
            <MapElement NoColumn="" Attribute="FechaInicioVigencia" Ignore="true" Default="01/01/2017" Format="" />
            <MapElement NoColumn="" Attribute="FechaFinVigencia" Ignore="true" Default="null" Format="" />
            <MapElement NoColumn="" Attribute="Version" Ignore="true" Default="1.0" Format="" />
          </MapElements>
        </Mapper>
      </Mappers>
      <Sheets>
        <Sheet NoSheet="1" SheetName="c_Aduana" RowNumberStartData="6" RowNumberStopData="54" Map="MapAduana"/>
      </Sheets>
    </ExcelMapping>
  </ExcelMappings>
</ExcelMappingSection>
	
```

Aquí hay un ejemplo del uso de multiples mappers para una hoja de Excel con distintos formatos:

```
<ExcelMapping Name="TipoComprobante">
      <Mappers>
        <Mapper Name="MapTipoComprobante">
          <MapElements>
            <MapElement NoColumn="0" Attribute="Clave" Ignore="false" Default="" Format=""/>
            <MapElement NoColumn="1" Attribute="Descripcion" Ignore="false" Default="" Format=""/>
            <MapElement NoColumn="2" Attribute="Valor_Maximo" Ignore="false" Default="" Format=""/>
            <MapElement NoColumn="4" Attribute="FechaInicioVigencia" Ignore="false" Default="" Format="" />
            <MapElement NoColumn="5" Attribute="FechaFinVigencia" Ignore="false" Default="" Format="" />
            <MapElement NoColumn="" Attribute="Version" Ignore="true" Default="1.0" Format="" />
          </MapElements>
        </Mapper>
        <Mapper Name="MapTipoComprobante2">
          <MapElements>
            <MapElement NoColumn="" Attribute="Clave" Ignore="true" Default="N" Format=""/>
            <MapElement NoColumn="1" Attribute="Descripcion" Ignore="false" Default="" Format=""/>
            <MapElement NoColumn="2" Attribute="Valor_Maximo_NS" Ignore="false" Default="" Format=""/>
            <MapElement NoColumn="3" Attribute="Valor_Maximo_NdS" Ignore="false" Default="" Format=""/>
            <MapElement NoColumn="4" Attribute="FechaInicioVigencia" Ignore="false" Default="" Format="" />
            <MapElement NoColumn="5" Attribute="FechaFinVigencia" Ignore="false" Default="" Format="" />
            <MapElement NoColumn="" Attribute="Version" Ignore="true" Default="1.0" Format="" />
          </MapElements>
        </Mapper>
      </Mappers>
      <Sheets>
        <Sheet NoSheet="1" SheetName="c_TipoDeComprobante" RowNumberStartData="6" RowNumberStopData = "8" Map="MapTipoComprobante"/>
        <Sheet NoSheet="2" SheetName="c_TipoDeComprobante" RowNumberStartData="10" RowNumberStopData = "10" Map="MapTipoComprobante2"/>
        <Sheet NoSheet="3" SheetName="c_TipoDeComprobante" RowNumberStartData="11" RowNumberStopData = "11" Map="MapTipoComprobante"/>
      </Sheets>
</ExcelMapping
```

Y otro ejemplo de un mapeo de multiples hojas hacia una misma colección con el mismo mapping y diferentes rangos de filas:

```
 <ExcelMapping Name="CodigoPostal">
      <Mappers>
        <Mapper Name="MapCP">
          <MapElements>
            <MapElement NoColumn="0" Attribute="CodigoPostal" Ignore="false" Default="" Format="00000"/>
            <MapElement NoColumn="1" Attribute="Estado" Ignore="false" Default="" Format="" />
            <MapElement NoColumn="2" Attribute="Municipio" Ignore="false" Default="" Format="000" />
            <MapElement NoColumn="3" Attribute="Localidad" Ignore="false" Default="" Format="00" />
            <MapElement NoColumn="" Attribute="FechaInicioVigencia" Ignore="true" Default="27/03/2017" Format="" />
            <MapElement NoColumn="" Attribute="FechaFinVigencia" Ignore="true" Default="null" Format="" />
            <MapElement NoColumn="" Attribute="Version" Ignore="true" Default="2.0" Format="" />
          </MapElements>
        </Mapper>
      </Mappers>
      <Sheets>
        <Sheet NoSheet="1" SheetName="c_CodigoPostal_Parte_1" RowNumberStartData="6" RowNumberStopData = "47889" Map="MapCP"/>
        <Sheet NoSheet="2" SheetName="c_CodigoPostal_Parte_2" RowNumberStartData="6" RowNumberStopData = "47898" Map="MapCP"/>
      </Sheets>
</ExcelMapping>
```

Hay que destacar que el formato se usa para forzar la conversión desde un campo que podria ser un número almacenado como texto, cosa común cuando no
se formatea correctamente la hoja de excel. Y necesitamos expresar numeros con ceros a la izquierda (ejem. **"0025"**).



## Ejemplos de Uso

El proyecto Web o de la aplicación donde se use la libreria debe contener la hoja de Excel como recurso "contenido" y con la caracteristica de "copiar siempre" o "copiar solo si es posterior" (ver las propiedades del archivo), para que se genere adecuadamente en el build o release.

```
ProyectoWeb/Aplicacion
 --|ExcelSources
    --|Archivo.xlsx

```

En el código de ejemplo se construye la clase como un repositorio que hereda de la clase  **BaseExcelReader< >** que es la que se encargará del 
trabajo de procesar los datos.

```
public class AduanaExcelReader : BaseExcelReader<AduanaResult>
{        
	public AduanaExcelReader() : base(
		"ExcelSourceDirectory"
	    "ExcelFileName.xls"
		"AduanaExcelMappingName")
	{           
            
	}

	public override void CalculateFields(ref AduanaResult obj)
	{            
		obj.Key = obj.Clave;
	}
}
```

Para usarlo basta con invocar el método **LoadData** y utilizar la propiedad **GetDataList** del ExcelReader.


```
 static void Main(string[] args)
{
	var reader = new AduanaExcelReader();
	reader.LoadData();
            
	foreach (var item in reader.GetDataList)
	{
		System.Console.WriteLine(item);
	}

	reader = null;
	
	//Probando la persistencia de los datos
	reader = new AduanaExcelReader();	
	var list = reader.GetDataList;
	
	System.Console.Read();
}
```

Al invocar **LoadData** una ConcurrentBag del tipo del expecificado se llenará y permanecerá en memoria hasta que termine el programa.
Por lo que no es necesario llamar al meétodo cada vez que se necesitan los datos del Reader.


## Calculate Fields

Para los casos en los que hay campos que dependen del valor de otros, existe este método.
**Se ejecuta siempre despues de se han llenado todos  los datos desde Excel en el objeto mapeado.**

En casi todas las implementaciones para este proyecto en particular se esta usando este método para asignar el campo **Key** de la clase
base, no todas las clases a mapear tienen el campo **Clave**, por lo que se decidio mapearlo para realizar las busquedas sobre el campo **Key** 
o realizar una implementación de la busqueda de manera especifica.

**Ejemplo:**
En el MonedaExcelReader, se están calculando 2 campos, el campo **Key** y el **PorcentajeVariacion** que en la hoja de catálogos de excel esta expresado
como un porcentaje (menor a 1), sin embargo en la implementación actual esta almacenado como un entero mayor que 100, por lo que estamos haciendo este ajuste
en el método para evitar errores en las validaciones.

```
public class MonedaExcelReader : BaseExcelReader<MonedaResult>
{
	public MonedaExcelReader() : base(ConfigurationHelper.ExcelSourceDirectory
		, ConfigurationHelper.ExcelSourceCatCfdiFileName
		, ConfigurationHelper.MonedaExcelMappingName)
	{
	}

	public override void CalculateFields(ref MonedaResult obj)
	{
		obj.Key = obj.Clave;
		if (obj.PorcentajeVariacion < 1)
		{
			obj.PorcentajeVariacion = obj.PorcentajeVariacion * 100;
		}
	}
}
```

## BaseRepository

Fue necesario implementar un repositorio similar al existente y usar la interface ICatalogRepository < T >
```
public class BaseRepository<T> : ICatalogRepository<T>
    where T : BaseResult, new()
{       
	public BaseRepository(string mainCacheKey = "", bool withCache = false)
	{
		MainCacheKey = mainCacheKey;
		WithCache = withCache;
	}

	public string MainCacheKey { get; set; }
        
	public bool WithCache { get; set; }
        
	public virtual bool GetByKey(string key)
	{
		return GetByKey(key, DateTime.Now);
	}

	public virtual bool GetByKey(string key, DateTime date)
	{ 
		return GetDataByKey(key, date) != null;
	}

	public virtual T GetDataByKey(string key)
	{
		return GetDataByKey(key, DateTime.Now);
	}

	public virtual T GetDataByKey(string key, DateTime date)
	{            
		if(WithCache)
			return SimpleCacheProvider<T>.Instance.GetCacheItem($"{MainCacheKey}:{key}", CachePopulate(key, date));

		return GetItem(key, date);
	}


	protected virtual T GetItem(string key, DateTime date)
	{
		if (string.IsNullOrEmpty(key)) return null;

		var list = ExcelReaderFactory.Create<T>().GetDataList.AsParallel().Where(t => t.Key == key);

		return list?.AsParallel().Where(delegate (T item)
			{
				if (item.FechaInicioVigencia != null)
				{
					if (DateTime.Compare(date, item.FechaInicioVigencia.Value) < 0)
					{
					return false;
					}
				}

				if (item.FechaFinVigencia != null)
				{
					return DateTime.Compare(date, item.FechaFinVigencia.Value) < 0;
				}

				return true;
			})
			.FirstOrDefault();
            
	}
	
	protected virtual Func<T> CachePopulate(string key, DateTime date)
	{
		return () => GetItem(key, date);
	}
}
```

El campo **WithCache** permite que la lista utilice **MemoryCache** para tener un acceso más rapido a los key-value mas usados.

Si **WithCache es igual a true** se utiliza la clase **SimpleCacheProvider** y el metodo **CachePopulate** (funcion anónima) para regresar el valor que 
necesita guardar la **MemoryCache**, es necesario tambíen establecer el campo **MainCacheKey** con la que se formará una llave de busqueda (**MainCacheKey:key**) en la **MemoryCache** y se pueda identificar adecuadamente este key-value adecuadamente.





