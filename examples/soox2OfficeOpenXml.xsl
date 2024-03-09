<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:soox="simplify-office-open-xml"
    xmlns:math="http://www.w3.org/2005/xpath-functions/math"
    xmlns:map="http://www.w3.org/2005/xpath-functions/map"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
    xmlns:my="my"
    exclude-result-prefixes="#all"
    extension-element-prefixes="soox"
    version="3.0">
    <xd:doc scope="stylesheet">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> Oct 6, 2021</xd:p>
            <xd:p><xd:b>Author:</xd:b>Olivier XILLO</xd:p>
            <xd:p>Demonstrate how to use SOOX to convert 
            SOOX document to Open Office document
            </xd:p>
        </xd:desc>
    </xd:doc>
    
    <!--xsl:output method="xml" standalone="yes" indent="yes"/-->
    
    <xsl:use-package name="soox" version="1.0"/>
    
    <xsl:output method="xml" indent="true"/>
    
    <xsl:variable name="base" select="static-base-uri()"/>
    
    <xsl:template name="xsl:initial-template">
        <xsl:variable name="my:workbook" select="doc('example-workbook.soox')"/>
        <xsl:variable name="xlsx-file" select="resolve-uri('generated-OfficeOpenXml-workbook.xlsx',$base)"/>
        <xsl:copy-of select="soox:toOfficeOpenXml( $my:workbook, map{'uri':$xlsx-file} )"/>
        <xsl:sequence select="soox:fromOfficeOpenXml( $xlsx-file, map{} )"/>
    </xsl:template>
    
    <xsl:template match="/">
        <xsl:variable name="xlsx-file" select="resolve-uri('generated-OfficeOpenXml-workbook.xlsx',$base)"/>
        <xsl:copy-of select="soox:toOfficeOpenXml( . , map{'uri':$xlsx-file} )"/>
        <xsl:sequence select="soox:fromOfficeOpenXml( $xlsx-file, map{} )"/>
    </xsl:template>
  
</xsl:stylesheet>