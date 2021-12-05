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
```
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
