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
            <xd:p>Generate the signature string for the cell fill style</xd:p>
        </xd:desc>
        <xd:param name="style">the cell style</xd:param>
        <xd:return>a string that uniquely identifies the fill style</xd:return>
    </xd:doc>
    <xsl:function name="soox:fill-signature" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <xsl:sequence select="$style/@fill-signature"/>
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
            <xsl:for-each-group select="($default-fill,$styles)" group-by="xs:string(@fill-signature)">
                <xsl:sort order="ascending"/>
                
                
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
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Returns the OOXML fill style index from a fill style signature</xd:p>
            <xd:p>The returned index is the position of the signature in the keys of the fill styles map</xd:p>
        </xd:desc>
        <xd:param name="signature">The fill style signature to find</xd:param>
        <xd:param name="styles">A fill styles map that associate a signature to a OOXML fill element</xd:param>
        <xd:return>The index of the matching signature or 0(=no fill) if not found</xd:return>
    </xd:doc>
    <xsl:function name="soox:fill-index-of" as="xs:integer">
        <xsl:param name="styles" as="map(xs:string,element(sml:fill))"/>
        <xsl:param name="signature" as="xs:string"/>
        
        <xsl:variable name="matching" as="xs:integer*"
            select="(map:keys($styles)[. ne $default-fill-signature])=>index-of($signature)"/>
        <xsl:sequence select="if($matching) then $matching[1] else 0"/>
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
        
        
        <!-- Generates a fills element containing fill elements; first element should be the default one -->
        <sml:fills count="{count(map:keys($fillsTablemap))}">
            
            <xsl:sequence select="$fillsTablemap($default-fill-signature)"/>
            <xsl:for-each select="map:keys($fillsTablemap)[. ne $default-fill-signature]">
                <xsl:sequence select="$fillsTablemap(current())"/>
            </xsl:for-each>
        </sml:fills>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>default fill specification</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:variable name="default-fill" as="element(s:style)">
        <s:style fill-style="none"/>
    </xsl:variable>
    
    <xsl:variable name="default-fill-signature" as="xs:string"
        select="'none'"/>
        
    
    <xd:doc>
        <xd:desc>
            <xd:p>valid values for @fill-style</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:variable name="fill-styles" as="xs:string+"
        select="('inherit','none','solid')"/>
    
</xsl:stylesheet>