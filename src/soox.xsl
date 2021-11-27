<?xml version="1.0" encoding="UTF-8"?>
<xsl:package name="soox" package-version="1.0"
    version="3.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
    xmlns:soox="simple-open-office-xml"
    exclude-result-prefixes="#all">
    
    <xd:doc scope="package">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> Nov 7, 2021</xd:p>
            <xd:p><xd:b>Author:</xd:b> Olivier XILLO</xd:p>
            <xd:p>The SOOX package provides a simple way to convert
                  Open Office document to/from simple XML representation.
                  Check SOOX website for a description of SOOX Xml format.
            </xd:p>
        </xd:desc>
    </xd:doc>
    
    <xsl:use-package name="soox:common" version="1.0"/>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Convert a SOOX document to an Open Office document</xd:p>
        </xd:desc>
        <xd:param name="doc">A SOOX document</xd:param>
        <xd:param name="options"></xd:param>
    </xd:doc>
    <xsl:function name="soox:toOpenOffice" visibility="public">
        <xsl:param name="doc" as="document-node()"/>
        <xsl:param name="options" as="map(*)"/>
        
        
        
        <xsl:choose>
            <xsl:when test="$doc/element()=>node-name() eq QName('soox','workbook')">
                <xsl:sequence select="soox:generate-OpenOffice-file-hierarchy( $doc, $options ) =>
                                      soox:zip-OpenOffice-file-hierarchy( resolve-uri($options?uri) )"/>        
            </xsl:when>
            <xsl:otherwise>
                <xsl:message>soox:toOpenOffice is not implemented for this element</xsl:message>        
            </xsl:otherwise>
        </xsl:choose>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Convert an Open Office document to a SOOX document</xd:p>
        </xd:desc>
        <xd:param name="oodoc">An Open Office document</xd:param>
        <xd:param name="options"></xd:param>
    </xd:doc>
    <xsl:function name="soox:fromOpenOffice" visibility="public">
        <xsl:param name="uri"/>
        <xsl:param name="options" as="map(*)"/>
        <xsl:sequence select="soox:unzip-OpenOffice-file-hierarchy( resolve-uri($uri) )"/>
    </xsl:function>
    
</xsl:package>