<?xml version="1.0" encoding="UTF-8"?>
<xsl:package name="soox" package-version="1.0"
    version="3.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
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
        
        
        
        <xsl:choose>
            <xsl:when test="$doc/element()=>node-name() eq QName('soox','workbook')">
                <xsl:sequence select="soox:generate-OfficeOpenXml-file-hierarchy( $doc, $options ) =>
                                      soox:zip-OfficeOpenXml-file-hierarchy( resolve-uri($options?uri) )"/>        
            </xsl:when>
            <xsl:otherwise>
                <xsl:message>soox:toOfficeOpenXml is not implemented for this element</xsl:message>        
            </xsl:otherwise>
        </xsl:choose>
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
        <xsl:sequence select="soox:unzip-OfficeOpenXml-file-hierarchy( resolve-uri($uri) )"/>
    </xsl:function>
    
</xsl:package>