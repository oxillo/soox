<?xml version="1.0" encoding="UTF-8"?>
<xsl:package name="soox" package-version="1.0"
    version="3.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
    xmlns:map="http://www.w3.org/2005/xpath-functions/map"
    xmlns:soox="simplify-office-open-xml"
    xmlns:opc="simplify-office-open-xml/open-packaging-conventions"
    exclude-result-prefixes="#all"
    extension-element-prefixes="opc">
    
    <xd:doc scope="package">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> Nov 7, 2021</xd:p>
            <xd:p><xd:b>Author:</xd:b> Olivier XILLO</xd:p>
            <xd:p>The SOOX package provides a simple way to convert
                  Office Open Xml document to/from simple XML representation.
                  Check SOOX website for a description of SOOX Xml format.
            </xd:p>
        </xd:desc>
    </xd:doc>
    
    <xsl:use-package name="soox:common" version="1.0"/>
    <!-- support for Open Packaging Conventions --> 
    <xsl:use-package name="soox:opc" version="1.0"/>
    <xsl:use-package name="soox:spreadsheet" version="1.0"/>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Convert a SOOX document to an Open Office document</xd:p>
        </xd:desc>
        <xd:param name="doc">A SOOX document</xd:param>
        <xd:param name="options"></xd:param>
    </xd:doc>
    <xsl:function name="soox:toOfficeOpenXml" visibility="public">
        <xsl:param name="doc" as="document-node()"/>
        <xsl:param name="options" as="map(*)"/>
        
        <xsl:variable name="soox-root-element" select="$doc/element()"/>
        <xsl:variable name="ooxml-parts" as="map(*)">
            <xsl:choose>
                <xsl:when test="$soox-root-element=>node-name() eq QName('soox','workbook')">
                    <xsl:sequence select="$soox-root-element=>soox:encode-spreadsheet_part()"/>        
                </xsl:when>
                <xsl:otherwise>
                    <xsl:message>soox:toOfficeOpenXml is not implemented for this element</xsl:message>        
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:sequence select="$ooxml-parts=>opc:zip(resolve-uri($options?uri))"/>              
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Convert an Office Open Xml document to a SOOX document</xd:p>
        </xd:desc>
        <xd:param name="oodoc">An Office Open Xml document</xd:param>
        <xd:param name="options"></xd:param>
    </xd:doc>
    <xsl:function name="soox:fromOfficeOpenXml" visibility="public">
        <xsl:param name="uri"/>
        <xsl:param name="options" as="map(*)"/>
        <xsl:variable name="ooxml-parts" as="map(*)">
            <xsl:variable name="unzipped" select="opc:unzip(resolve-uri($uri))"/>
            <xsl:map>
                <xsl:for-each select="map:keys($unzipped)">
                    <xsl:map-entry key="'/'||current()" select="$unzipped(current())"/>
                </xsl:for-each>
            </xsl:map>
        </xsl:variable> 
        
        
        <xsl:if test="$ooxml-parts=>map:contains('/[Content_Types].xml')">
            <!-- decode [Content_Types].xml -->
            
            <!--xsl:variable name="content-types" select="$ooxml-parts => soox:extract-xmlfile-from-file-hierarchy('[Content_Types].xml')"/>
            <xsl:variable name="workbook" select="$content-types/Types/Override[@ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml']/@PartName" xpath-default-namespace="http://schemas.openxmlformats.org/package/2006/content-types"/>
            <xsl:variable name="document" select="$content-types/Types/Override[@ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml']/@PartName" xpath-default-namespace="http://schemas.openxmlformats.org/package/2006/content-types"/>
            <xsl:if test="$workbook"-->
                <!--xsl:sequence select="$ooxml-parts => soox:extract-office-part() => soox:decode-spreadsheet-part()"/-->
            <xsl:sequence select="$ooxml-parts =>  soox:decode-spreadsheet-part('/xl/workbook.xml')"/>
            <!--/xsl:if>
            <xsl:if test="$document">
                <xsl:message expand-text="true">Document is {$document}</xsl:message>
            </xsl:if-->
        </xsl:if>
    </xsl:function>
    
</xsl:package>