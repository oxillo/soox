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
            <xd:p><xd:b>Created on:</xd:b> Nov 18, 2023</xd:p>
            <xd:p><xd:b>Author:</xd:b> Olivier</xd:p>
            <xd:p>Helper functions for font manipulation</xd:p>
        </xd:desc>
    </xd:doc>
    
    <xsl:include href="utils_colors.xsl"/>
    
    
    
    <xd:doc>
        <xd:desc>
            <xd:p></xd:p>
        </xd:desc>
        <xd:param name="inherited"></xd:param>
        <xd:param name="local"></xd:param>
    </xd:doc>
    <xsl:function name="soox:cascade-font-style">
        <xsl:param name="inherited" as="map(*)"/>
        <xsl:param name="local" as="element(s:style)"/>
        
        
        <xsl:variable name="from-detailed" as="map(*)">
            <xsl:map>
                <xsl:apply-templates mode="soox:spreadsheet-styles-cascade" select="(
                    $local/@font-family,$local/@font-size,$local/@font-style,$local/@font-weight)"/>
            </xsl:map>
        </xsl:variable>
        <xsl:sequence select="map:merge(($inherited,$from-detailed),map{'duplicates':'use-last'})"/>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p></xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:template match="@font-family" mode="soox:spreadsheet-styles-cascade">
        <xsl:map-entry key="local-name()" select="."/>
    </xsl:template>
    
    <xsl:template match="@font-size" mode="soox:spreadsheet-styles-cascade">
        <xsl:map-entry key="local-name()" select="."/>
    </xsl:template>
    
    <xsl:template match="@font-style" mode="soox:spreadsheet-styles-cascade">
        <xsl:if test="data()=('normal','italic')">
            <xsl:map-entry key="local-name()" select="."/>    
        </xsl:if>
        <!--xsl:message expand-text="true">Invalid font style specification &quot;{$style/@font-style}&quot;</xsl:message-->
    </xsl:template>
    
    <xsl:template match="@font-weight" mode="soox:spreadsheet-styles-cascade">
        <xsl:map-entry key="local-name()" select="."/>
    </xsl:template>
    
    
    
    
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Build a map to associates the font signature to the generated OOXML font element</xd:p>
            <xd:p>The map does not contains the 'none' font style</xd:p>
        </xd:desc>
        <xd:param name="styles">SOOX cell styles</xd:param>
        <xd:return>A map that associates the font signature with the generated OOXML font element</xd:return>
    </xd:doc>
    <xsl:function name="soox:buildFontStyleMap" as="map(xs:string,element(sml:font))">
        <xsl:param name="styles" as="element(s:style)*"/>
        
        <xsl:map>
            <xsl:for-each-group select="$styles" group-by="xs:string(@font-signature)">
                <xsl:map-entry key="current-grouping-key()">
                    <!-- Tokenize the font-signature to get in this order : family, size, weight and style -->
                    <xsl:variable name="font-style" select="current-grouping-key()=>tokenize('#')"/>
                    <sml:font>
                        <sml:sz val="{$font-style[2]}"/>
                        <!--sml:color rgb="FF000000"/-->
                        <sml:name val="{$font-style[1]}"/>
                        <!--sml:family val="2"/>
                        <sml:charset val="238"/-->
                        <xsl:if test="$font-style[3] = 'bold'">
                            <sml:b/>
                        </xsl:if>
                        <xsl:choose>
                            <xsl:when test="$font-style[4] = 'italic'">
                                <sml:i/>
                            </xsl:when>
                            <xsl:otherwise/>
                        </xsl:choose>
                    </sml:font>
                </xsl:map-entry>
            </xsl:for-each-group>
        </xsl:map>
    </xsl:function>
    
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Generates a table for all unique font styles</xd:p>
        </xd:desc>
        <xd:param name="fontsTablemap">a map of sml:font</xd:param>
        <xd:return>a font table inside a fonts element</xd:return>
    </xd:doc>
    <xsl:function name="soox:fonts-table" as="element(sml:fonts)">
        <xsl:param name="fontsTablemap" as="map(xs:string,element(sml:font))"/>
        
        <!-- Generates a fonts element containing font elements; first element should be the default one -->
        <sml:fonts count="{count(map:keys($fontsTablemap))}">
            <xsl:for-each select="map:keys($fontsTablemap)">
                <xsl:sort order="ascending" stable="yes"/>
                <xsl:sequence select="$fontsTablemap(current())"/>
            </xsl:for-each>
        </sml:fonts>
    </xsl:function>
   
   
</xsl:stylesheet>