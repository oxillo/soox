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
            <xd:p>Generates a 'style' element with detailed font attributes from a style with the shortened 'border' attrinbute</xd:p>
            <xd:p><![CDATA[<style font="12 'Arial'"/> generates <style font-family="Arial" font-size="12"/>]]></xd:p>
        </xd:desc>
        <xd:param name="style">A 'style' element</xd:param>
        <xd:return>A 'style' element</xd:return>
    </xd:doc>
    <xsl:function name="soox:expand-border-attributes" as="element(s:style)">
        <xsl:param name="style" as="element(s:style)"/>
        
        <!--TODO-->
        <s:style/>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Return the left border using by priority order : 
            <xd:ul>
                <xd:li>the 'border-left-*' attributes value</xd:li>
                <xd:li>the values from in the 'border-left' attribute</xd:li>
                <xd:li>the values from in the 'border' attribute</xd:li>
                <xd:li>the default values (no border, width=1, color=black)</xd:li>
            </xd:ul></xd:p>
        </xd:desc>
        <xd:param name="style">A SOOX 'style' element</xd:param>
        <xd:return>A string that is the border specification</xd:return>
    </xd:doc>
    <xsl:function name="soox:border-left" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <xsl:variable name="border-style" as="xs:string"
            select="($style/@border-left-style,soox:expand-border-attributes($style)/@border-left-style,'none')[1]"/>
        <xsl:choose>
            <xsl:when test="$border-style='none'">none</xsl:when>
            <xsl:otherwise>
                <xsl:variable name="border-color" as="xs:string"
                    select="($style/@border-left-color,soox:expand-border-attributes($style)/@border-left-color,'black')[1]=>soox:parseColor()"/>
                <xsl:sequence select="string-join(($border-style,$border-color),'#')"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Return the right border using by priority order : 
                <xd:ul>
                    <xd:li>the 'border-right-*' attributes value</xd:li>
                    <xd:li>the values from in the 'border-right' attribute</xd:li>
                    <xd:li>the values from in the 'border' attribute</xd:li>
                    <xd:li>the default values (no border, width=1, color=black)</xd:li>
                </xd:ul></xd:p>
        </xd:desc>
        <xd:param name="style">A SOOX 'style' element</xd:param>
        <xd:return>A string that is the border specification</xd:return>
    </xd:doc>
    <xsl:function name="soox:border-right" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <xsl:variable name="border-style" as="xs:string"
            select="($style/@border-right-style,soox:expand-border-attributes($style)/@border-right-style,'none')[1]"/>
        <xsl:choose>
            <xsl:when test="$border-style='none'">none</xsl:when>
            <xsl:otherwise>
                <xsl:variable name="border-color" as="xs:string"
                    select="($style/@border-right-color,soox:expand-border-attributes($style)/@border-right-color,'black')[1]=>soox:parseColor()"/>
                <xsl:sequence select="string-join(($border-style,$border-color),'#')"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Return the top border using by priority order : 
                <xd:ul>
                    <xd:li>the 'border-top-*' attributes value</xd:li>
                    <xd:li>the values from in the 'border-top' attribute</xd:li>
                    <xd:li>the values from in the 'border' attribute</xd:li>
                    <xd:li>the default values (no border, width=1, color=black)</xd:li>
                </xd:ul></xd:p>
        </xd:desc>
        <xd:param name="style">A SOOX 'style' element</xd:param>
        <xd:return>A string that is the border specification</xd:return>
    </xd:doc>
    <xsl:function name="soox:border-top" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <xsl:variable name="border-style" as="xs:string"
            select="($style/@border-top-style,soox:expand-border-attributes($style)/@border-top-style,'none')[1]"/>
        <xsl:choose>
            <xsl:when test="$border-style='none'">none</xsl:when>
            <xsl:otherwise>
                <xsl:variable name="border-color" as="xs:string"
                    select="($style/@border-top-color,soox:expand-border-attributes($style)/@border-top-color,'black')[1]=>soox:parseColor()"/>
                <xsl:sequence select="string-join(($border-style,$border-color),'#')"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Return the bottom border using by priority order : 
                <xd:ul>
                    <xd:li>the 'border-bottom-*' attributes value</xd:li>
                    <xd:li>the values from in the 'border-bottom' attribute</xd:li>
                    <xd:li>the values from in the 'border' attribute</xd:li>
                    <xd:li>the default values (no border, width=1, color=black)</xd:li>
                </xd:ul></xd:p>
        </xd:desc>
        <xd:param name="style">A SOOX 'style' element</xd:param>
        <xd:return>A string that is the border specification</xd:return>
    </xd:doc>
    <xsl:function name="soox:border-bottom" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <xsl:variable name="border-style" as="xs:string"
            select="($style/@border-bottom-style,soox:expand-border-attributes($style)/@border-bottom-style,'none')[1]"/>
        <xsl:choose>
            <xsl:when test="$border-style='none'">none</xsl:when>
            <xsl:otherwise>
                <xsl:variable name="border-color" as="xs:string"
                    select="($style/@border-bottom-color,soox:expand-border-attributes($style)/@border-bottom-color,'black')[1]=>soox:parseColor()"/>
                <xsl:sequence select="string-join(($border-style,$border-color),'#')"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Generate the signature string for the cell borders style</xd:p>
        </xd:desc>
        <xd:param name="style">the cell style</xd:param>
        <xd:return>a string that uniquely identifies the borders style</xd:return>
    </xd:doc>
    <xsl:function name="soox:border-signature" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <xsl:sequence select="if ($style) then 
                map{
                'left':soox:border-left($style),
                'right':soox:border-right($style),
                'top':soox:border-top($style),
                'bottom':soox:border-bottom($style)
                }=>serialize(map{'method':'adaptive'}) 
                else soox:border-signature($default-border)"/>
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
            <xsl:for-each-group select="($default-border,$styles)" group-by="soox:border-signature(.)">
                <xsl:map-entry key="current-grouping-key()">
                    <xsl:variable name="border-style" select="current-group()[1]"/>
                    <sml:border>
                        <sml:left>
                            <xsl:variable name="args" select="soox:border-left($border-style)=>tokenize('#')"/>
                            <xsl:if test="$args[1] ne 'none'">
                                <xsl:attribute name="style" select="$args[1]"/>
                                <sml:color rgb="{$args[2]}"/>
                            </xsl:if>
                        </sml:left>
                        <sml:right>
                            <xsl:variable name="args" select="soox:border-right($border-style)=>tokenize('#')"/>
                            <xsl:if test="$args[1] ne 'none'">
                                <xsl:attribute name="style" select="$args[1]"/>
                                <sml:color rgb="{$args[2]}"/>
                            </xsl:if>
                        </sml:right>
                        <sml:top>
                            <xsl:variable name="args" select="soox:border-top($border-style)=>tokenize('#')"/>
                            <xsl:if test="$args[1] ne 'none'">
                                <xsl:attribute name="style" select="$args[1]"/>
                                <sml:color rgb="{$args[2]}"/>
                            </xsl:if>
                        </sml:top>
                        <sml:bottom>
                            <xsl:variable name="args" select="soox:border-bottom($border-style)=>tokenize('#')"/>
                            <xsl:if test="$args[1] ne 'none'">
                                <xsl:attribute name="style" select="$args[1]"/>
                                <sml:color rgb="{$args[2]}"/>
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
            <xd:p>The default font specification</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:variable name="default-border" as="element(s:style)">
        <s:style border-left-style="none" border-right-style="none" border-top-style="none" border-bottom-style="none"/>
    </xsl:variable>
    
</xsl:stylesheet>