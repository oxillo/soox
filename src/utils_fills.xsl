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
            <xd:p><xd:b>Created on:</xd:b> Dec 05, 2023</xd:p>
            <xd:p><xd:b>Author:</xd:b>XILLO Olivier</xd:p>
            <xd:p>Helper functions for fill style</xd:p>
        </xd:desc>
    </xd:doc>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p></xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:template match="@fill-style" mode="soox:spreadsheet-styles-cascade">
        <xsl:map-entry key="'fill-style'" select="."/>
        <xsl:if test="data()='none'">
            <xsl:map-entry key="'fill-color'" select="soox:parse-color('black','invalid')"/>
        </xsl:if>
    </xsl:template>
    
    <xsl:template match="@fill-color" mode="soox:spreadsheet-styles-cascade">
        <xsl:variable name="color" select="soox:parse-color(.,'invalid')"/>
        <xsl:if test="$color ne 'invalid'">
           <xsl:map-entry key="'fill-style'">solid</xsl:map-entry>
           <xsl:map-entry key="'fill-color'" select="$color"/>
        </xsl:if>
    </xsl:template>
    
    <xsl:template match="@fill" mode="soox:spreadsheet-styles-cascade">
        <xsl:variable name="tokens" select="normalize-space()=>tokenize(' ')"/>
        <xsl:map-entry key="'fill-style'" select="$tokens[1]"/>
        <xsl:if test="$tokens[2] and $tokens[1] ne 'none'">
            <xsl:variable name="color" select="soox:parse-color(.,'invalid')"/>
            <xsl:if test="$color ne 'invalid'">
                <xsl:map-entry key="'fill-color'" select="$color"/>
            </xsl:if>
        </xsl:if>
        <xsl:if test="$tokens[1] = 'none'">
            <xsl:map-entry key="'fill-color'" select="soox:parse-color('black','invalid')"/>
        </xsl:if>
    </xsl:template>
    
    
    <xsl:function name="soox:cascade-fill-style">
        <xsl:param name="inherited" as="map(*)"/>
        <xsl:param name="local" as="element(s:style)"/>
        
        
        <xsl:variable name="from-fill" as="map(*)*">
            <xsl:map>
                <xsl:apply-templates select="$local/@fill" mode="soox:spreadsheet-styles-cascade"/>
            </xsl:map>
        </xsl:variable>
        <xsl:variable name="from-fill-color" as="map(*)*">
            <xsl:map>
                <xsl:apply-templates select="$local/@fill-color" mode="soox:spreadsheet-styles-cascade"/>
            </xsl:map>
        </xsl:variable>
        <xsl:variable name="from-fill-style" as="map(*)">
            <xsl:map>
                <xsl:apply-templates select="$local/@fill-style" mode="soox:spreadsheet-styles-cascade"/>
            </xsl:map>
        </xsl:variable>
        <xsl:sequence select="map:merge(($inherited,$from-fill,$from-fill-color,$from-fill-style),map{'duplicates':'use-last'})"/>
    </xsl:function>
    
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Build a map to associates the fill signature to the generated OOXML fill element</xd:p>
            <xd:p>The map does not contains the 'none' fill style</xd:p>
        </xd:desc>
        <xd:param name="styles">SOOX cell styles</xd:param>
        <xd:return>A map that associates the fill signature with the generated OOXML fill element</xd:return>
    </xd:doc>
    <xsl:function name="soox:buildFillStyleMap" as="map(xs:string,element(sml:fill))">
        <xsl:param name="styles" as="element(s:style)*"/>
        
        <xsl:map>
            <!-- get the fill items from their unique signature -->
            <xsl:for-each-group select="($styles)" group-by="xs:string(@fill-signature)">
                
                <xsl:map-entry key="current-grouping-key()">
                    <xsl:variable name="tokens" select="current-grouping-key()=>tokenize('#')"/>
                    <xsl:choose>
                        <xsl:when test="$tokens[1] ne 'none'">
                            <sml:fill>
                                <sml:patternFill patternType="solid">
                                    <sml:fgColor rgb="{$tokens[2]}"/>
                                    <sml:bgColor rgb="{$tokens[2]}"/>
                                </sml:patternFill>
                            </sml:fill>
                        </xsl:when>
                        <xsl:otherwise>
                            <sml:fill>
                                <sml:patternFill patternType="none"/>
                            </sml:fill>
                                    
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:map-entry>
            </xsl:for-each-group>
        </xsl:map>
    </xsl:function>
    
    
    
    <xsl:function name="soox:index-of" as="xs:integer">
        <xsl:param name="styles" as="map(xs:string,element())"/>
        <xsl:param name="signature" as="xs:string"/>
        
        <xsl:sequence select="map:keys($styles)=>index-of($signature) - 1"/>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Generates a table for all unique border styles</xd:p>
        </xd:desc>
        <xd:param name="cellStyles">a sequence of all the cell styles</xd:param>
        <xd:return>a border table inside a borders element</xd:return>
    </xd:doc>
    <xsl:function name="soox:fills-table" as="element(sml:fills)">
        <xsl:param name="fillsTablemap" as="map(xs:string,element(sml:fill))"/>
        
        
        <!-- Generates a fills element containing fill elements -->
        <sml:fills count="{count(map:keys($fillsTablemap))}">
            <xsl:for-each select="map:keys($fillsTablemap)">
                <xsl:sort order="ascending" stable="yes"/>
                <xsl:sequence select="$fillsTablemap(current())"/>
            </xsl:for-each>
        </sml:fills>
    </xsl:function>
    
        
    
    <xd:doc>
        <xd:desc>
            <xd:p>valid values for @fill-style</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:variable name="fill-styles" as="xs:string+"
        select="('inherit','none','solid')"/>
    
</xsl:stylesheet>