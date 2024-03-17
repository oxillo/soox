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
            <xd:p></xd:p>
        </xd:desc>
        <xd:param name="inherited"></xd:param>
        <xd:param name="local"></xd:param>
    </xd:doc>
    <xsl:function name="soox:cascade-border-style">
        <xsl:param name="inherited" as="map(*)"/>
        <xsl:param name="local" as="element(s:style)"/>
        
        
        <xsl:variable name="from-border" as="map(*)*">
            <xsl:apply-templates select="$local/@border" mode="soox:spreadsheet-styles-cascade"/>  
        </xsl:variable>
        <xsl:variable name="from-border-stylecolor" as="map(*)*">
            <xsl:apply-templates select="($local/@border-color,$local/@border-style)" mode="soox:spreadsheet-styles-cascade"/>
        </xsl:variable>
        <xsl:variable name="from-border-detailed" as="map(*)">
            <xsl:map>
                <xsl:apply-templates mode="soox:spreadsheet-styles-cascade" select="(
                    $local/@border-left-style,$local/@border-right-style,$local/@border-top-style,$local/@border-bottom-style,
                    $local/@border-left-color,$local/@border-right-color,$local/@border-top-color,$local/@border-bottom-color)"/>
            </xsl:map>
        </xsl:variable>
        <xsl:sequence select="map:merge(($inherited,$from-border,$from-border-stylecolor,$from-border-detailed),map{'duplicates':'use-last'})"/>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Expand @border to a border styles map</xd:p>
            <xd:p>The first word of @border attribute is the style of the border ('thin','none',...)</xd:p>
            <xd:p>The second optional word of @border attribute is the color of the border ('green','blue',...)</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:template match="@border" mode="soox:spreadsheet-styles-cascade">
        <xsl:variable name="tokens" select="normalize-space(.)=>tokenize(' ')"/>
        
        <xsl:map>
            <xsl:if test="$tokens[1]=$borders-styles">
                <xsl:map-entry key="'border-left-style'" select="$tokens[1]"/>
                <xsl:map-entry key="'border-right-style'" select="$tokens[1]"/>
                <xsl:map-entry key="'border-top-style'" select="$tokens[1]"/>
                <xsl:map-entry key="'border-bottom-style'" select="$tokens[1]"/>
            </xsl:if>
            
            <xsl:if test="$tokens[2]">
                <xsl:variable name="parsed-color" select="soox:parse-color($tokens[2],'invalid')"/>
                <xsl:if test="$parsed-color ne 'invalid'">
                    <xsl:map-entry key="'border-left-color'" select="$parsed-color"/>
                    <xsl:map-entry key="'border-right-color'" select="$parsed-color"/>
                    <xsl:map-entry key="'border-top-color'" select="$parsed-color"/>
                    <xsl:map-entry key="'border-bottom-color'" select="$parsed-color"/>
                </xsl:if>
            </xsl:if>
        </xsl:map>
    </xsl:template>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Expand @border-color to a border color map</xd:p>
            <xd:p>A single value defines the color for all 4 borders</xd:p>
            <xd:p>2 values define color for top/bottom and left/right</xd:p>
            <xd:p>3 values define color for top, left/right and bottom</xd:p>
            <xd:p>4 values define color for top, right, bottom and left</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:template match="@border-color" mode="soox:spreadsheet-styles-cascade">
        <xsl:variable name="tokens" select="(normalize-space(.)=>tokenize(' '))!soox:parse-color(.,'invalid')"/>
        
        <xsl:map>
            <xsl:choose>
                <xsl:when test="count($tokens)=1">
                    <xsl:if test="$tokens[1] ne 'invalid'">
                        <xsl:map-entry key="'border-left-color'" select="$tokens[1]"/>
                        <xsl:map-entry key="'border-right-color'" select="$tokens[1]"/>
                        <xsl:map-entry key="'border-top-color'" select="$tokens[1]"/>
                        <xsl:map-entry key="'border-bottom-color'" select="$tokens[1]"/>
                    </xsl:if>
                </xsl:when>
                <xsl:when test="count($tokens)=2">
                    <xsl:if test="$tokens[1] ne 'invalid'">
                        <xsl:map-entry key="'border-top-color'" select="$tokens[1]"/>
                        <xsl:map-entry key="'border-bottom-color'" select="$tokens[1]"/>
                    </xsl:if>
                    <xsl:if test="$tokens[2] ne 'invalid'">
                        <xsl:map-entry key="'border-left-color'" select="$tokens[2]"/>
                        <xsl:map-entry key="'border-right-color'" select="$tokens[2]"/>
                    </xsl:if>
                </xsl:when>
                <xsl:when test="count($tokens)=3">
                    <xsl:if test="$tokens[1] ne 'invalid'">
                        <xsl:map-entry key="'border-top-color'" select="$tokens[1]"/>
                    </xsl:if>
                    <xsl:if test="$tokens[2] ne 'invalid'">
                        <xsl:map-entry key="'border-left-color'" select="$tokens[2]"/>
                        <xsl:map-entry key="'border-right-color'" select="$tokens[2]"/>
                    </xsl:if>
                    <xsl:if test="$tokens[3] ne 'invalid'">
                        <xsl:map-entry key="'border-bottom-color'" select="$tokens[3]"/>
                    </xsl:if>
                </xsl:when>
                <xsl:when test="count($tokens)=4">
                    <xsl:if test="$tokens[1] ne 'invalid'">
                        <xsl:map-entry key="'border-top-color'" select="$tokens[1]"/>
                    </xsl:if>
                    <xsl:if test="$tokens[2] ne 'invalid'">
                        <xsl:map-entry key="'border-right-color'" select="$tokens[2]"/>
                    </xsl:if>
                    <xsl:if test="$tokens[3] ne 'invalid'">
                        <xsl:map-entry key="'border-bottom-color'" select="$tokens[3]"/>
                    </xsl:if>
                    <xsl:if test="$tokens[4] ne 'invalid'">
                        <xsl:map-entry key="'border-left-color'" select="$tokens[4]"/>
                    </xsl:if>
                </xsl:when>
                <xsl:otherwise/>
            </xsl:choose>
        </xsl:map>
    </xsl:template>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Expand @border-style to a border style map</xd:p>
            <xd:p>A single value defines the style for all 4 borders</xd:p>
            <xd:p>2 values define style for top/bottom and left/right</xd:p>
            <xd:p>3 values define style for top, left/right and bottom</xd:p>
            <xd:p>4 values define style for top, right, bottom and left</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:template match="@border-style" mode="soox:spreadsheet-styles-cascade">
        <xsl:variable name="tokens" select="normalize-space(.)=>tokenize(' ')"/>
        
        <xsl:map>
            <xsl:choose>
                <xsl:when test="count($tokens)=1">
                    <xsl:if test="$tokens[1]=$borders-styles">
                        <xsl:map-entry key="'border-left-style'" select="$tokens[1]"/>
                        <xsl:map-entry key="'border-right-style'" select="$tokens[1]"/>
                        <xsl:map-entry key="'border-top-style'" select="$tokens[1]"/>
                        <xsl:map-entry key="'border-bottom-style'" select="$tokens[1]"/>
                    </xsl:if>
                </xsl:when>
                <xsl:when test="count($tokens)=2">
                    <xsl:if test="$tokens[1]=$borders-styles">
                        <xsl:map-entry key="'border-top-style'" select="$tokens[1]"/>
                        <xsl:map-entry key="'border-bottom-style'" select="$tokens[1]"/>
                    </xsl:if>
                    <xsl:if test="$tokens[2]=$borders-styles">
                        <xsl:map-entry key="'border-left-style'" select="$tokens[2]"/>
                        <xsl:map-entry key="'border-right-style'" select="$tokens[2]"/>
                    </xsl:if>
                </xsl:when>
                <xsl:when test="count($tokens)=3">
                    <xsl:if test="$tokens[1]=$borders-styles">
                        <xsl:map-entry key="'border-top-style'" select="$tokens[1]"/>
                    </xsl:if>
                    <xsl:if test="$tokens[2]=$borders-styles">
                        <xsl:map-entry key="'border-left-style'" select="$tokens[2]"/>
                        <xsl:map-entry key="'border-right-style'" select="$tokens[2]"/>
                    </xsl:if>
                    <xsl:if test="$tokens[3]=$borders-styles">
                        <xsl:map-entry key="'border-bottom-style'" select="$tokens[3]"/>
                    </xsl:if>
                </xsl:when>
                <xsl:when test="count($tokens)=4">
                    <xsl:if test="$tokens[1]=$borders-styles">
                        <xsl:map-entry key="'border-top-style'" select="$tokens[1]"/>
                    </xsl:if>
                    <xsl:if test="$tokens[2]=$borders-styles">
                        <xsl:map-entry key="'border-right-style'" select="$tokens[2]"/>
                    </xsl:if>
                    <xsl:if test="$tokens[3]=$borders-styles">
                        <xsl:map-entry key="'border-bottom-style'" select="$tokens[3]"/>
                    </xsl:if>
                    <xsl:if test="$tokens[4]=$borders-styles">
                        <xsl:map-entry key="'border-left-style'" select="$tokens[4]"/>
                    </xsl:if>
                </xsl:when>
                <xsl:otherwise/>
            </xsl:choose>
        </xsl:map>
    </xsl:template>
    
    <xd:doc>
        <xd:desc>
            <xd:p></xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:template match="@border-left-style|@border-right-style|@border-top-style|@border-bottom-style" mode="soox:spreadsheet-styles-cascade">
        <xsl:map-entry key="local-name()" select="."/>
    </xsl:template>
    
    <xd:doc>
        <xd:desc>
            <xd:p></xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:template match="@border-left-color|@border-right-color|@border-top-color|@border-bottom-color" mode="soox:spreadsheet-styles-cascade">
        <xsl:variable name="parsed-color" select="soox:parse-color(data(),'invalid')"/>
        <xsl:if test="$parsed-color ne 'invalid'">
            <xsl:map-entry key="local-name()" select="$parsed-color"/>  
        </xsl:if>
    </xsl:template>
    
    
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Build a map to associates the border signature to the generated OOXML border element</xd:p>
        </xd:desc>
        <xd:param name="styles">SOOX cell styles</xd:param>
        <xd:return>A map that associates the border signature with the generated OOXML font element</xd:return>
    </xd:doc>
    <xsl:function name="soox:buildBorderStyleMap" as="xs:string*">
        <xsl:param name="styles" as="element(s:style)*"/>
        
        <xsl:sequence select="distinct-values($styles/@border-signature)=>sort()"/>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Generates a table for all unique border styles</xd:p>
        </xd:desc>
        <xd:param name="bordersTablemap">a sequence of all the cell styles</xd:param>
        <xd:return>a border table inside a borders element</xd:return>
    </xd:doc>
    <xsl:function name="soox:borders-table" as="element(sml:borders)">
        <xsl:param name="signatures" as="xs:string*"/>
        
        <!-- Generates a borders element containing border elements; first element should be the default one -->
        <sml:borders count="{count($signatures)}">
            <xsl:for-each select="$signatures">
                <!-- Tokenize the border-signature to get in this order : [1]left-style, left-color, [3]right-style, right-color,
                        [5]top-style,top-color, [7]bottom-style, bottom-color -->
                <xsl:variable name="border-style" select="current()=>tokenize('#')"/>
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
            </xsl:for-each>
        </sml:borders>
    </xsl:function>
    
    
    <xsl:variable name="borders-styles" as="xs:string+"
        select="('inherit','none','thin','medium','dashed','dotted','thick','double','hair',
        'mediumDashed','dashDot','mediumDashDot','dashDotDot','mediumDashDotDot','slantDashDot')"/>
    
    <!--xsl:variable name="borders" select="('border-left-style','border-left-color',
        'border-right-style','border-right-color',
        'border-top-style','border-top-color',
        'border-bottom-style','border-bottom-color')"/-->
    
</xsl:stylesheet>