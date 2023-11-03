# Simplify Office Open XML

When it comes to exhange spreadsheets, documents or slides, Office Open Xml files are generally well accepted by end-users.
However, when it comes to their use in automation, they are not so friendly due to the fact that the data is spread in several files.
This is where SOOX comes to the rescue. 
It can convert an Open Office file to a single XML document (and vice-versa).
SOOX do not plan to support all the features available in Office Open Xml standard.
It will concentrate on the main content - cell values for spreadsheet, paragraph text for workprocessor.

The project is still in a very early stage. Consider it as unstable !
Our first target is Spreadsheets.


## Usage

SOOX has only 2 functions 
`soox:toOfficeOpenXml( soox-document, options-map )` : Converts the SOOX document to the Office Open Xml file hierarchy.
Use 'uri' key in options-map to write the Office Open Xml document to that uri

`soox:fromOfficeOpenXml( uri, options-map )` : Converts the Office Open Xml document in URI to a SOOX document

```xslt
        <xsl:copy-of select="soox:toOfficeOpenXml( $sooxworkbook, map{'uri':resolve-uri('officeopenXml-workbook.xlsx')} )"/>
        <xsl:sequence select="soox:fromOfficeOpenXml( resolve-uri('officeopenxml-workbook.xlsx'), map{} )"/>
```


## Simplify Open Office XML format

### Spreadsheet
```xml
<workbook xmln="soox">
    <worksheet name="1st sheet">
        <data>
            <cell col="1" row="1">
                <v>Value A1</v>
            </cell> 
            <cell col="2" row="3">
                <v>Value B3</v>
            </cell>
        </data>
    </worksheet>
    <worksheet name="2nd sheet">
        <data>
            <cell col="1" row="1">
                <v>1.0</v>
            </cell> 
            <cell col="2" row="3">
                <v>2.0</v>
            </cell>
        </data>
    </worksheet>    
</workbook>    
```

### Styling
 
Columns and rows can be customized by adding a `<style/>` element inside the `<worksheet/>` element.
The column width is set with an `width` element with the following attributes :
  - The `col` attribute (mandatory) corresponds to the index of the column (starting at 1).
  - The `w` attribute (exclusive with `px`) corresponds to the width of the column.
  - The `px` attribute (exclusive with `w`) corresponds to the pixel width of the column.

The row height is set with an `height` element with the following attributes :
  - The `row` attribute (mandatory) corresponds to the index of the row (starting at 1).
  - The `h` attribute (mandatory) corresponds to the height of the row. The format may specify a unit ('cm' for centimeters, 'px' for pixels)

```xml
<worksheet name="1st sheet">
    <style>
        <width col="1" w="15.3"/>
        <width col="2" px="300"/>
        <height row="1" h="300 px"/>
        <height row="2" h="1cm"/>
        <height row="3" h="20"/>
    </style>
    <data>
      ...
    </data>
</worksheet>
```
