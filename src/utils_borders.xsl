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
            <xd:p><xd:b>Created on:</xd:b> Nov 20, 2023</xd:p>
            <xd:p><xd:b>Author:</xd:b> Olivier</xd:p>
            <xd:p>Helper functions for borders manipulation</xd:p>
        </xd:desc>
    </xd:doc>
    
    
    
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Generate the signature string for the cell borders style</xd:p>
        </xd:desc>
        <xd:param name="style">the cell style</xd:param>
        <xd:return>a string that uniquely identifies the borders style</xd:return>
    </xd:doc>
    <xsl:function name="soox:border-signature" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <xsl:sequence select="$style/@border-signature"/>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Build a map to associates the border signature to the generated OOXML border element</xd:p>
        </xd:desc>
        <xd:param name="styles">SOOX cell styles</xd:param>
        <xd:return>A map that associates the border signature with the generated OOXML font element</xd:return>
    </xd:doc>
    <xsl:function name="soox:buildBorderStyleMap" as="map(xs:string,element(sml:border))">
        <xsl:param name="styles" as="element(s:style)*"/>
        
        <xsl:map>
            <xsl:for-each-group select="($default-border,$styles)" group-by="xs:string(@border-signature)">
                <xsl:map-entry key="current-grouping-key()">
                    <!-- Tokenize the border-signature to get in this order : [1]left-style, left-color, [3]right-style, right-color,
                        [5]top-style,top-color, [7]bottom-style, bottom-color -->
                    <xsl:variable name="border-style" select="current-grouping-key()=>tokenize('#')"/>
                    
                   <sml:border>
                        <sml:left>
                            <xsl:if test="$border-style[1] ne 'none'">
                                <xsl:attribute name="style" select="$border-style[1]"/>
                                <sml:color rgb="{$border-style[2]}"/>
                            </xsl:if>
                        </sml:left>
                        <sml:right>
                            <xsl:if test="$border-style[3] ne 'none'">
                                <xsl:attribute name="style" select="$border-style[3]"/>
                                <sml:color rgb="{$border-style[4]}"/>
                            </xsl:if>
                        </sml:right>
                        <sml:top>
                            <xsl:if test="$border-style[5] ne 'none'">
                                <xsl:attribute name="style" select="$border-style[5]"/>
                                <sml:color rgb="{$border-style[6]}"/>
                            </xsl:if>
                        </sml:top>
                        <sml:bottom>
                            <xsl:if test="$border-style[7] ne 'none'">
                                <xsl:attribute name="style" select="$border-style[7]"/>
                                <sml:color rgb="{$border-style[8]}"/>
                            </xsl:if>
                        </sml:bottom>
                    </sml:border>
                </xsl:map-entry>
            </xsl:for-each-group>
        </xsl:map>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Returns the OOXML fill style index from a font style signature</xd:p>
            <xd:p>The returned index is the position of the signature in the keys of the font styles map</xd:p>
        </xd:desc>
        <xd:param name="styles">A font styles map that associate a signature to a OOXML font element</xd:param>
        <xd:param name="signature">The font style signature to find</xd:param>
        <xd:return>The index of the matching signature or 0(=default font) if not found</xd:return>
    </xd:doc>
    <xsl:function name="soox:border-index-of" as="xs:integer">
        <xsl:param name="styles" as="map(xs:string,element(sml:border))"/>
        <xsl:param name="signature" as="xs:string"/>
        
        <xsl:variable name="matching" as="xs:integer*"
            select="(map:keys($styles)[. ne $default-border=>soox:border-signature()])=>index-of($signature)"/>
        <xsl:sequence select="if($matching) then $matching[1] else 0"/>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Generates a table for all unique border styles</xd:p>
        </xd:desc>
        <xd:param name="cellStyles">a sequence of all the cell styles</xd:param>
        <xd:return>a border table inside a borders element</xd:return>
    </xd:doc>
    <xsl:function name="soox:borders-table" as="element(sml:borders)">
        <xsl:param name="bordersTablemap" as="map(xs:string,element(sml:border))"/>
        
        <!-- Generates a borders element containing border elements; first element should be the default one -->
        <sml:borders count="{1 + count(map:keys($bordersTablemap))}">
            
            <xsl:variable name="default-border-signature" select="$default-border/@border-signature"/>
            <xsl:sequence select="$bordersTablemap($default-border-signature)"/>
            <xsl:for-each select="map:keys($bordersTablemap)[. ne $default-border-signature]">
                <xsl:sequence select="$bordersTablemap(current())"/>
            </xsl:for-each>
        </sml:borders>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>The default font specification</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:variable name="default-border" as="element(s:style)">
        <s:style border-left-style="none" border-right-style="none" border-top-style="none" border-bottom-style="none"
            border-left-color="black" border-right-color="black" border-top-color="black" border-bottom-color="black" border-signature="none#black#none#black#none#black#none#black"/>
    </xsl:variable>
    
    <xsl:variable name="borders-styles" as="xs:string+"
        select="('inherit','none','thin','medium','dashed','dotted','thick','double','hair',
        'mediumDashed','dashDot','mediumDashDot','dashDotDot','mediumDashDotDot','slantDashDot')"/>
    
    <xsl:variable name="borders" select="('border-left-style','border-left-color',
        'border-right-style','border-right-color',
        'border-top-style','border-top-color',
        'border-bottom-style','border-bottom-color')"/>
    
</xsl:stylesheet>