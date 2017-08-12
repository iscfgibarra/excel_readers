#ExcelReaders

##Tabla de Contenido

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


###ExcelMapping

Un ExcelMapping puede contener múltiples mappers y sheets. 

* **Name**: El nombre con que se identifica, este nombre permite asociar la configuración con el ExcelReader especifico.

###Mapper

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



## Ejemplos de Uso

El proyecto Web o de la aplicación donde se use la libreria debe contener la hoja de Excel como recurso "contenido" y con la caracteristica de "copiar siempre" o "copiar solo si es posterior" (ver las propiedades del archivo), para que se genere adecuadamente en el build o release.

```
ProyectoWeb/Aplicacion
 --|ExcelSources
    --|Archivo.xlsx

```

En el código de ejemplo se construye 

## Calculate Fields




