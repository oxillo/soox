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
    
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Propagate the numeric-format style from upper elements</xd:p>
        </xd:desc>
        <xd:param name="inherited"></xd:param>
        <xd:param name="local"></xd:param>
    </xd:doc>
    <xsl:function name="soox:cascade-numeric-format">
        <xsl:param name="inherited" as="map(*)"/>
        <xsl:param name="local" as="element(s:style)"/>
        
        <xsl:if test="$local/@numeric-format">
            <xsl:variable name="from-local" as="map(*)">
                <xsl:map>
                    <xsl:map-entry key="'numeric-format'" select="$local/@numeric-format"/>
                </xsl:map>
            </xsl:variable>
            <xsl:variable name="cascaded-style" select="map:merge(($inherited,$from-local),map{'duplicates':'use-last'})"/>
            <xsl:sequence select="$cascaded-style=>map:put('numeric-format-signature',$cascaded-style('numeric-format'))"/>
        </xsl:if>
        <xsl:if test="not($local/@numeric-format)">
            <xsl:sequence select="$inherited"/>
        </xsl:if>
        
    </xsl:function>
    
    
    
    
    <xd:doc>
        <xd:desc>
            <xd:p></xd:p>
        </xd:desc>
        <xd:param name="numericFormatTableMap"></xd:param>
        <xd:return></xd:return>
    </xd:doc>
    <xsl:function name="soox:numeric-formats-table" as="element(sml:numFmts)?">
        <xsl:param name="numericFormatTableMap" as="map(xs:string,xs:integer)"/>
        
        <!-- build the list of custom formats -->
        <xsl:variable name="numFmts" as="element(sml:numFmt)*">
            <xsl:for-each select="map:keys($numericFormatTableMap)[not(. = map:keys($soox:builtin-formatCodes))]">
                <sml:numFmt numFmtId="{$numericFormatTableMap(current())}" formatCode="{current()}"/>
            </xsl:for-each>
        </xsl:variable>
        
        <!-- If custom numeric formats are defined, export them -->
        <xsl:if test="count($numFmts) gt 0">
            <sml:numFmts count="{count($numFmts)}">
                <xsl:sequence select="$numFmts"/>
            </sml:numFmts>    
        </xsl:if>        
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Build the table that maps the numeric format to it index</xd:p>
            <xd:p>the table includes builtin numeric formats</xd:p>
        </xd:desc>
        <xd:param name="styles">The cell styles</xd:param>
        <xd:return></xd:return>
    </xd:doc>
    <xsl:function name="soox:buildNumFmtMap" as="map(xs:string,xs:integer)">
        <xsl:param name="styles" as="element(s:style)*"/>
        
        <!-- create the map for user defined numeric formats-->
        <xsl:variable name="extra-fmtCode" as="map(xs:string,xs:integer)">
            <xsl:variable name="signatures" as="xs:string*" 
                select="distinct-values($styles/@numeric-format-signature)"/>
            <xsl:map>
                <xsl:for-each select="$signatures[not(.=map:keys($soox:builtin-formatCodes))]">
                    <!-- define an index (above 164) to avoid builtin formatCode -->
                    <xsl:map-entry key="current()" select="163 + position()"/>
                </xsl:for-each>
            </xsl:map>
        </xsl:variable>
        <!-- merge with builtin numeric formats -->
        <xsl:sequence select="map:merge(($soox:builtin-formatCodes,$extra-fmtCode))"/>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Predefined format codes defined by the ECMA 376 standard</xd:p>
            <xd:p>Warning : some international codes have been ignored</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:variable name="soox:builtin-formatCodes" as="map(xs:string,xs:integer)"
        static="yes" visibility="private" select="map{
        'General':0,
        '0':1,
        '0.00':2,
        '#,##0':3,
        '#,##0.00':4,
        '0%':9,
        '0.00%':10,
        '0.00E+00':11,
        '# ?/?':12,
        '# ??/??':13,
        'mm-dd-yy':14,
        'd-mmm-yy':15,
        'd-mmm':16,
        'mmm-yy':17,
        'h:mm AM/PM':18,
        'h:mm:ss AM/PM':19,
        'h:mm':20,
        'h:mm:ss':21,
        'm/d/yy h:mm':22,
        '#,##0 ;(#,##0':37,
        '#,##0 ;[Red](#,##0)':38,
        '#,##0.00;(#,##0.00)':39,
        '#,##0.00;[Red](#,##0.00)':40,
        'mm:ss':45,
        '[h]:mm:ss':46,
        'mmss.0':47,
        '##0.0E+0':48,
        '@':49
        }"/>
    
</xsl:stylesheet>