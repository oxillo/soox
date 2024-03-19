<?xml version="1.0" encoding="UTF-8"?>
<xsl:package
    name="soox:spreadsheet"
    package-version="1.0"
    version="3.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
    xmlns:soox="simplify-office-open-xml"
    xmlns:s="soox"
    xmlns:map="http://www.w3.org/2005/xpath-functions/map"
    xmlns:sml="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships"
    exclude-result-prefixes="#all"
    extension-element-prefixes="soox">
    
    <xd:doc scope="package">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> Nov 7, 2021</xd:p>
            <xd:p><xd:b>Author:</xd:b> Olivier XILLO</xd:p>
            <xd:p></xd:p>
        </xd:desc>
    </xd:doc>

    <xsl:use-package name="soox:utils" version="1.0"/>
    
    <xsl:include href="spreadsheet-styles.xsl"/>
    <xsl:include href="spreadsheet-templates.xsl"/>
    
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Decode the specified spreadsheet</xd:p>
        </xd:desc>
        <xd:param name="parts">the parts in the package</xd:param>
        <xd:param name="partname">the name of the part to decode</xd:param>
    </xd:doc>
    <xsl:function name="soox:decode-spreadsheet-part" visibility="public">
        <xsl:param name="parts"/>
        <xsl:param name="partname"/>
        
        <!-- Extract workbook file -->
        <xsl:variable name="workbook" select="$parts => soox:get-content( $partname )"/>
        
        <!-- Decompose partname into segments to extract the base and basename -->
        <!-- /xl/workbook.xml will give base='/xl/' and basename='workbook.xml' --> 
        <xsl:variable name="segments" select="tokenize($partname,'/')"/>
        <xsl:variable name="base" select="string-join($segments[position() lt last()],'/')||'/'"/>
        <xsl:variable name="basename" select="$segments[last()]"/>
        
        <!-- Extract relationship file for partname -->
        <xsl:variable name="relationship" select="$parts => soox:get-content($base||'_rels/'||$basename||'.rels')"/>
        
        <!-- Extract list of shared-strings -->
        <xsl:variable name="sharedstrings-list" as="xs:string*">
            <xsl:variable name="sharedstrings">
                <xsl:variable name="fname" select="$relationship => soox:get-relationships-of-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings')"/>
                <xsl:sequence select="$parts => soox:get-content($base||$fname[1]/@Target )"/>
            </xsl:variable>
            <xsl:sequence select="$sharedstrings//*:t/text()"/>
        </xsl:variable>
        
        <!-- Extract styles -->
        <xsl:variable name="styles">
            <xsl:variable name="fname" select="$relationship => soox:get-relationships-of-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles')"/>
            <xsl:sequence select="$parts => soox:get-content($base||$fname[1]/@Target )"/>
        </xsl:variable>
        
        <!-- Process workbook -->
        <xsl:apply-templates select="$workbook" mode="soox:fromOfficeOpenXml">
            <xsl:with-param name="file-hierarchy" select="$parts" tunnel="yes"/>
            <xsl:with-param name="sharedstrings-list" select="$sharedstrings-list" tunnel="yes"/>
            <xsl:with-param name="relationship" select="$relationship" tunnel="yes"/>
            <xsl:with-param name="base" select="$base" tunnel="yes"/>
        </xsl:apply-templates>
    </xsl:function>
    
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Encode a soox element as a spreadsheet</xd:p>
        </xd:desc>
        <xd:param name="elem">the element to encode</xd:param>
        <xd:param name="base">the base path in the package where to store generated parts</xd:param>
    </xd:doc>
    <xsl:function name="soox:encode-spreadsheet_part" visibility="public">
        <xsl:param name="elem"/>
        <xsl:param name="base"/>
        
        <!-- do style cascading in order that every cell has *all* style attributes (ex: @border-style expanded to @border-*-style -->
        <xsl:variable name="wbk" as="element()">
            <xsl:apply-templates select="$elem" mode="soox:spreadsheet-styles-cascade">
                <xsl:with-param name="inherited-style" tunnel="yes" select="$default-cell-style"/>
            </xsl:apply-templates>
        </xsl:variable>
        
        <!-- Compute items shared across worksheets : shared strings and styles -->
        <xsl:variable name="all-cell-values" select="$wbk//s:cell/s:v/text()[not(. castable as xs:date or . castable as xs:double) ]"/>
        <xsl:variable name="shared-strings" as="xs:string*" select="$all-cell-values=>distinct-values()"/>
        <xsl:variable name="shared-strings-map" as="map(xs:string,xs:integer)">
            <xsl:map>
                <xsl:for-each select="$shared-strings">
                    <xsl:map-entry key="current()" select="position()"/>
                </xsl:for-each>
            </xsl:map>
        </xsl:variable>
        <xsl:variable name="sharedStrings.xml" select="soox:sharedStrings.xml($shared-strings,count($all-cell-values))"/>
        <xsl:variable name="styles.xml" select="$wbk => soox:styles.xml()"/>
        <xsl:variable name="cell-styles-map" select="$wbk//s:cell/s:style => soox:buildCellStylesMap()"/>
        
        <xsl:variable name="worksheets" as="map(*)">
            <xsl:variable name="ws" as="map(*)*">
                <xsl:variable name="wstype" select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'"/>
                <xsl:for-each select="$wbk/s:worksheet">
                    <xsl:variable name="wsfilename" select="$base||'worksheets/sheet'||position()||'.xml'"/>
                    <xsl:sequence select="map:entry($wsfilename,current()=>soox:worksheet.xml($shared-strings-map, $cell-styles-map))"/>
                </xsl:for-each>
            </xsl:variable>
            <xsl:sequence select="map:merge($ws)"/>
        </xsl:variable>
        <xsl:variable name="files" select="map:merge((
            $worksheets,
            map:entry($base||'styles.xml', $styles.xml),
            map:entry($base||'sharedStrings.xml', $sharedStrings.xml)))"/>
        <xsl:variable name="workbook.xml.rels" select="$files => soox:build-relationship-file('xl')"/>
        <xsl:variable name="workbook.xml" select="$wbk => soox:workbook.xml( $workbook.xml.rels )"/>
        
        
        <xsl:sequence select="map:merge(($files,map{$base||'workbook.xml':$workbook.xml, $base||'_rels/workbook.xml.rels':$workbook.xml.rels}))"/>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Encode a soox element as a spreadsheet and store parts under '/xl/' path</xd:p>
        </xd:desc>
        <xd:param name="elem">the element to encode</xd:param>
    </xd:doc>
    <xsl:function name="soox:encode-spreadsheet_part" visibility="public">
        <xsl:param name="elem"/>
        <xsl:sequence select="soox:encode-spreadsheet_part($elem,'/xl/')"/>
    </xsl:function>
    
    
    <xsl:function name="soox:workbook.xml" visibility="private">
        <xsl:param name="simple_workbook"/>
        <xsl:param name="relationship-file"/>
        
        <xsl:variable name="relationships"
            select="$relationship-file('content')//*:Relationship[@Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet']"/>
        
        <xsl:variable name="content">
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                mc:Ignorable="x15"
                xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
                
                <fileVersion appName="Calc"/>
                <workbookPr backupFile="false" showObjects="all" date1904="false"/>
                <workbookProtection/>
                <bookViews>
                    <workbookView showHorizontalScroll="true" showVerticalScroll="true" showSheetTabs="true"
                        xWindow="0" yWindow="0" windowWidth="16384" windowHeight="8192" tabRatio="500"
                        firstSheet="0" activeTab="1"/>
                </bookViews>
                
                
                
                
                <!--fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"/>
        <workbookPr defaultThemeVersion="153222"/>
        <mc:AlternateContent
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <mc:Choice Requires="x15">
            <x15ac:absPath url="C:\Users\to24007\Dev\XSLT\Excel\"
              xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"
            />
          </mc:Choice>
        </mc:AlternateContent>
        <bookViews>
          <workbookView xWindow="0" yWindow="0" windowWidth="24300" windowHeight="14250"/>
        </bookViews-->
                <sheets>
                    <xsl:for-each select="$simple_workbook/s:worksheet">
                        <xsl:variable name="index" select="position()"/>
                        <xsl:variable name="rId" select="$relationships[@Target=>ends-with($index||'.xml')]/@Id"/>
                        <sheet name="{@name}" sheetId="{$index}" state="visible" r:id="{$rId}"/>
                    </xsl:for-each>
                </sheets>
                <calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.0001"/>
                <extLst>
                    <ext xmlns:loext="http://schemas.libreoffice.org/"
                        uri="{{7626C862-2A13-11E5-B345-FEFF819CDC9F}}">
                        <loext:extCalcPr stringRefSyntax="ExcelA1"/>
                    </ext>
                </extLst>
                
                <!--calcPr calcId="152511"/>
        <extLst>
          <ext uri="{{140A7094-0E35-4892-8432-C4D2E57EDEB5}}"
            xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
            <x15:workbookPr chartTrackingRefBase="1"/>
          </ext>
        </extLst-->
            </workbook>
        </xsl:variable>
        
        <xsl:sequence select="map{
            'content': $content,
            'content-type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
            'relationship-type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
            }"/>
        
    </xsl:function>
    
    
    
    
    <xd:doc>
        <xd:desc>
            <xd:p></xd:p>
        </xd:desc>
        <xd:param name="simple_worksheet"></xd:param>
    </xd:doc>
    <xsl:function name="soox:worksheet.xml" visibility="private">
        <xsl:param name="simple_worksheet"/>
        <xsl:param name="shared-strings" as="map(xs:string,xs:integer)"/>
        <xsl:param name="cell-styles-map" as="map(xs:string,xs:integer)"/>
        
        <xsl:variable name="content">
            <xsl:apply-templates select="$simple_worksheet" mode="soox:toOfficeOpenXml">
                <xsl:with-param name="shared-strings" select="$shared-strings" tunnel="yes"/>
                <xsl:with-param name="cell-styles-map" select="$cell-styles-map" tunnel="yes"/>
            </xsl:apply-templates>
        </xsl:variable>
        <xsl:sequence select="map{
            'content': $content,
            'content-type':'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
            'relationship-type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
            }"/>
    </xsl:function>
    
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Generates the sharedStrings file</xd:p>
        </xd:desc>
        <xd:param name="simple_workbook"></xd:param>
    </xd:doc>
    <xsl:function name="soox:sharedStrings.xml" visibility="private">
        <xsl:param name="shared-strings" as="xs:string*"/>
        <xsl:param name="all-strings-count" as="xs:integer"/>
        
        <xsl:variable name="content">
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <xsl:attribute name="count" select="$all-strings-count"/>
                <xsl:attribute name="uniqueCount" select="$shared-strings => count()"/>
                <xsl:for-each select="$shared-strings">
                    <si>
                        <t>
                            <xsl:attribute name="xml:space">preserve</xsl:attribute>
                            <xsl:value-of select="current()"/>
                        </t>
                    </si>
                </xsl:for-each>
            </sst>
        </xsl:variable>
        
        <xsl:sequence select="map{
            'content': $content,
            'content-type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
            'relationship-type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
            }"/>
    </xsl:function>
    
    
    
        
</xsl:package>