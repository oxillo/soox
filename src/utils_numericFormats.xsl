<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:map="http://www.w3.org/2005/xpath-functions/map"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
    xmlns:sml="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:soox="simplify-office-open-xml"
    xmlns:s="soox"
    extension-element-prefixes="soox"
    exclude-result-prefixes="#all"
    expand-text="yes"
    version="3.0">
    <xd:doc scope="stylesheet">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> Dec 16, 2023</xd:p>
            <xd:p><xd:b>Author:</xd:b> Olivier</xd:p>
            <xd:p>Helper functions for numeric formats</xd:p>
        </xd:desc>
    </xd:doc>
    
    <xsl:function name="soox:numeric-formats-table" as="element(sml:numFmts)?">
        <xsl:param name="cellStyles" as="element(s:style)*"/>
        
        <!-- build the list of custom formats -->
        <xsl:variable name="numFmts" as="element(sml:numFmt)*">
            <xsl:variable name="m" select="soox:buildNumFmtMap($cellStyles)"/>
            <xsl:for-each select="map:keys($m)">
                <xsl:if test="not($m(current()) = map:keys($soox:default-numeric-formats))">
                    <sml:numFmt numFmtId="{$m(current())}" formatCode="{current()}"/>
                </xsl:if>
            </xsl:for-each>
        </xsl:variable>
        
        <!-- If custom numeric formats are defined, export them -->
        <xsl:if test="count($numFmts) gt 0">
            <sml:numFmts count="{count($numFmts)}">
                <xsl:sequence select="$numFmts"/>
            </sml:numFmts>    
        </xsl:if>
        
    </xsl:function>
    
    
    <xsl:function name="soox:buildNumFmtMap" as="map(xs:string,xs:integer)">
        <xsl:param name="styles" as="element(s:style)*"/>
        
        <xsl:variable name="reversed-default-numeric-formats" as="map(xs:string,xs:integer)">        
            <xsl:map>
                <xsl:for-each select="map:keys($soox:default-numeric-formats)">
                    <xsl:map-entry key="$soox:default-numeric-formats(current())" select="current()"/>
                </xsl:for-each>
            </xsl:map>
        </xsl:variable>
        
        <xsl:variable name="numeric-formats" as="xs:string*"
            select="distinct-values(for $s in $styles return $s/@numeric-format)"/>
        <xsl:map>
            <xsl:sequence select="$reversed-default-numeric-formats"/>
            <xsl:for-each select="$numeric-formats[not(. = map:keys($reversed-default-numeric-formats))]">
                <xsl:map-entry key="current()" select="position() + 163"/>
            </xsl:for-each>    
        </xsl:map>
        
        
    </xsl:function>
    
    <xsl:function name="soox:numeric-format-index-of" as="xs:integer">
        <xsl:param name="map" as="map(xs:string,xs:integer)"/>
        <xsl:param name="signature" as="xs:string"/>
        
        <xsl:sequence select="($map($signature),0)[1]"/>
    </xsl:function>
    
    <xsl:function name="soox:numeric-format-signature" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <xsl:sequence select="($style/@numeric-format,'General')[1]"/>
    </xsl:function>
    
    <xsl:variable name="soox:default-numeric-formats" as="map(xs:integer,xs:string)"
        select="map{
        0:'General',
        1:'0',
        2:'0.00',
        3:'#,##0',
        4:'#,##0.00',
        9:'0%',
        10:'0.00%',
        11:'0.00E+00',
        12:'# ?/?',
        13:'# ??/??',
        14:'mm-dd-yy',
        15:'d-mmm-yy',
        16:'d-mmm',
        17:'mmm-yy',
        18:'h:mm AM/PM',
        19:'h:mm:ss AM/PM',
        20:'h:mm',
        21:'h:mm:ss',
        22:'m/d/yy h:mm',
        37:'#,##0 ;(#,##0',
        38:'#,##0 ;[Red](#,##0)',
        39:'#,##0.00;(#,##0.00)',
        40:'#,##0.00;[Red](#,##0.00)',
        45:'mm:ss',
        46:'[h]:mm:ss',
        47:'mmss.0',
        48:'##0.0E+0',
        49:'@'
        }"/>
    
</xsl:stylesheet>