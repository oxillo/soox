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
            <xd:p>Generate the signature string for the cell fill style</xd:p>
        </xd:desc>
        <xd:param name="style">the cell style</xd:param>
        <xd:return>a string that uniquely identifies the fill style</xd:return>
    </xd:doc>
    <xsl:function name="soox:fill-signature" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <xsl:variable name="fill-tokens" select="$style/@fill=>normalize-space()=>tokenize(' ')"/>
        <xsl:variable name="solid-fill" as="xs:string?"
            select="if ($style/@fill-color and not($style/@fill-style)) then 'solid' else ()"/>
        <xsl:variable name="fill-style" select="($style/@fill-style, $solid-fill, $fill-tokens[1],'none')[1]" as="xs:string"/>
        <xsl:variable name="fill-color" select="if ($fill-style = 'none') then () else ($style/@fill-color, $fill-tokens[2],'black')[1]" as="xs:string*"/>
        <xsl:choose>
            <xsl:when test="$fill-style ne 'none'">
                <xsl:sequence select="$fill-style||'#'||soox:parse-color(($style/@fill-color, $fill-tokens[2],'black')[1])"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:sequence select="$fill-style"/>
            </xsl:otherwise>
        </xsl:choose>
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
            <xsl:for-each-group select="($default-fill,$styles)" group-by="soox:fill-signature(.)">
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
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Generates a table for all unique border styles</xd:p>
        </xd:desc>
        <xd:param name="cellStyles">a sequence of all the cell styles</xd:param>
        <xd:return>a border table inside a borders element</xd:return>
    </xd:doc>
    <xsl:function name="soox:fills-table" as="element(sml:fills)">
        <xsl:param name="cellStyles" as="element(s:style)*"/>
        
        <!-- Generates a map {"fill-signature": sml:fill element} -->
        <xsl:variable name="fillsTablemap" as="map(xs:string,element(sml:fill))"
            select="soox:buildFillStyleMap($cellStyles)"/>
        
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
        select="$default-fill=>soox:fill-signature()"/>
        
    
    <xd:doc>
        <xd:desc>
            <xd:p>valid values for @fill-style</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:variable name="fill-styles" as="xs:string+"
        select="('inherit','none','solid')"/>
    
</xsl:stylesheet>