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
    xmlns:arch="http://expath.org/ns/archive"
    xmlns:file="http://expath.org/ns/file"
    xmlns:bin="http://expath.org/ns/binary"
    xmlns:sml="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships"
    exclude-result-prefixes="#all"
    extension-element-prefixes="soox arch file bin">
    
    <xd:doc scope="package">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> Nov 7, 2021</xd:p>
            <xd:p><xd:b>Author:</xd:b> Olivier XILLO</xd:p>
            <xd:p></xd:p>
        </xd:desc>
    </xd:doc>
    
    <!--xsl:use-package name="soox:common" version="1.0"/-->
    <!--xsl:use-package name="soox:utils" version="1.0"/-->


    <xsl:include href="spreadsheet-styles.xsl"/>
    <xsl:include href="spreadsheet-templates.xsl"/>
    
    
    
    
    <xsl:function name="soox:decode-spreadsheet-part" visibility="public">
        <xsl:param name="file-hierarchy"/>
        
        <xsl:variable name="office-fname" select="$file-hierarchy('$root')('name')"/>
        <xsl:variable name="relationship" select="$file-hierarchy => soox:extract-xmlfile-from-file-hierarchy('_rels/'||$office-fname||'.rels')"/>
        <xsl:variable name="workbook" select="$file-hierarchy => soox:extract-xmlfile-from-file-hierarchy( $office-fname )"/>
        <xsl:variable name="sharedstrings">
            <xsl:variable name="fname" select="$relationship => soox:get-relationships-of-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings')"/>
            <xsl:sequence select="$file-hierarchy => soox:extract-xmlfile-from-file-hierarchy( $fname[1]/@Target )"/>
        </xsl:variable>
        <xsl:variable name="sharedstrings-list" select="$sharedstrings//*:t/text()"/>
        <xsl:variable name="styles">
            <xsl:variable name="fname" select="$relationship => soox:get-relationships-of-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles')"/>
            <xsl:sequence select="$file-hierarchy => soox:extract-xmlfile-from-file-hierarchy( $fname[1]/@Target )"/>
        </xsl:variable>
        
        <xsl:apply-templates select="$workbook" mode="soox:fromOfficeOpenXml">
            <xsl:with-param name="file-hierarchy" select="$file-hierarchy" tunnel="yes"/>
            <xsl:with-param name="sharedstrings-list" select="$sharedstrings-list" tunnel="yes"/>
            <xsl:with-param name="relationship" select="$relationship" tunnel="yes"/>
        </xsl:apply-templates>
    </xsl:function>
    
    
    
    <xsl:function name="soox:encode-spreadsheet_part" visibility="public">
        <xsl:param name="elem"/>
        
        <xsl:variable name="base" select="'/xl'"/>
        <xsl:variable name="sharedStrings.xml" select="$elem => soox:sharedStrings.xml()"/>
        <xsl:variable name="worksheets" as="map(*)">
            <xsl:variable name="ws" as="map(*)*">
                <xsl:variable name="wstype" select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'"/>
                <xsl:variable name="shared-strings" select="$sharedStrings.xml('content')//*:sst/*:si/*:t/text()"/>
                <!--application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml-->
                <xsl:for-each select="$elem/s:worksheet">
                    <xsl:variable name="wsfilename" select="$base||'/worksheets/sheet'||position()||'.xml'"/>
                    <!--xsl:sequence select="current()=>xlwfile:worksheet.xml()=>xlwfn:pack-into($wsfilename, $wstype)"/-->
                    <xsl:sequence select="map:entry($wsfilename,current()=>soox:worksheet.xml($shared-strings))"/>
                </xsl:for-each>
            </xsl:variable>
            <xsl:sequence select="map:merge($ws)"/>
        </xsl:variable>
        <xsl:variable name="files" select="map:merge((
            $worksheets,
            $elem => soox:styles.xml($base||'/styles.xml'),
            map:entry($base||'/sharedStrings.xml', $sharedStrings.xml)))"/>
        <xsl:variable name="workbook.xml.rels" select="$files => soox:build-relationship-file('xl')"/>
        <xsl:variable name="workbook.xml" select="$elem => soox:workbook.xml( $workbook.xml.rels )"/>
        
        
        <xsl:sequence select="map:merge(($files,map{$base||'/workbook.xml':$workbook.xml, $base||'/_rels/workbook.xml.rels':$workbook.xml.rels}))"/>
    </xsl:function>
    
    
    <xsl:function name="soox:workbook.xml" visibility="public">
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
    <xsl:function name="soox:worksheet.xml" visibility="public">
        <xsl:param name="simple_worksheet"/>
        <xsl:param name="shared-strings" as="xs:string*"/>
        
        <xsl:variable name="content">
            <xsl:apply-templates select="$simple_worksheet" mode="soox:toOfficeOpenXml">
                <xsl:with-param name="shared-strings" select="$shared-strings" tunnel="yes"/>
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
    <xsl:function name="soox:sharedStrings.xml" visibility="public">
        <xsl:param name="simple_workbook"/>
        
        <!-- Collect the strings -->
        <xsl:variable name="all-cell-values" select="$simple_workbook//s:cell/s:v/text()[not(. castable as xs:date or . castable as xs:double) ]"/>
        <xsl:variable name="distinct-cell-values" select="$all-cell-values => distinct-values()"/>
        
        <xsl:variable name="content">
            <sst  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <xsl:attribute name="count" select="$all-cell-values => count()"/>
                <xsl:attribute name="uniqueCount" select="$distinct-cell-values => count()"/>
                <xsl:for-each select="$distinct-cell-values">
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