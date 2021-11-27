# Simple Open Office XML

When it comes to exhange spreadsheets, documents or slides, Open Office files are generally well accepted by end-users.
However, when it comes to their use in automation, they are not so friendly due to the fact that the data is spread in several files.
This is where SOOX comes to the rescue. 
It can convert an Open Office file to a single XML document (and vice-versa).
SOOX do not plan to support all the features available in Open Office standard.
It will concentrate on the main content - cell values for spreadsheet, paragraph text for workprocessor.

The project is still in a very early stage. Consider it as unstable !
Our first target is Spreadsheets.


## Usage

SOOX has only 2 functions 
`soox:toOpenOffice( soox-document, options-map )` : Converts the SOOX document to the Open Office file hierarchy.
Use 'uri' key in options-map to write the Open Office document to that uri

`soox:fromOpenOffice( uri, options-map )` : Converts the Open Office document in URI to a SOOX document

```xslt
        <xsl:copy-of select="soox:toOpenOffice( $sooxworkbook, map{'uri':resolve-uri('openoffice-workbook.xlsx')} )"/>
        <xsl:sequence select="soox:fromOpenOffice( resolve-uri('openoffice-workbook.xlsx'), map{} )"/>
```


## Simple Open Office XML format

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
