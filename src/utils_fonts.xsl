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
            <xd:p>Generates a 'style' element with detailed font attributes from a style with the shortened 'font' attrinbute</xd:p>
            <xd:p><![CDATA[<style font="12 'Arial'"/> generates <style font-family="Arial" font-size="12"/>]]></xd:p>
        </xd:desc>
        <xd:param name="style">A 'style' element</xd:param>
        <xd:return>A 'style' element</xd:return>
    </xd:doc>
    <xsl:function name="soox:font-from-short-attribute" as="element(s:style)">
        <xsl:param name="style" as="element(s:style)"/>
        
        <!--TODO-->
        <s:style/>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Return the font-family using by priority order : 
            <xd:ul>
                <xd:li>the 'font-family' attribute value</xd:li>
                <xd:li>the family specified in the 'font' attribute</xd:li>
                <xd:li>the default font-family</xd:li>
            </xd:ul></xd:p>
        </xd:desc>
        <xd:param name="style">A SOOX 'style' element</xd:param>
        <xd:return>A string that is the font-family</xd:return>
    </xd:doc>
    <xsl:function name="soox:get-font-family" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
                
        <xsl:sequence select="($style/@font-family,soox:font-from-short-attribute($style)/@font-family,$default-font/@font-family)[1]"/>
    </xsl:function>
        
    <xd:doc>
        <xd:desc>
            <xd:p>Return the font-size using by priority order : 
                <xd:ul>
                    <xd:li>the 'font-size' attribute value</xd:li>
                    <xd:li>the size specified in the 'font' attribute</xd:li>
                    <xd:li>the default font-size</xd:li>
                </xd:ul></xd:p>
        </xd:desc>
        <xd:param name="style">A SOOX 'style' element</xd:param>
        <xd:return>A string that is the font-size</xd:return>
    </xd:doc>
    <xsl:function name="soox:get-font-size" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <xsl:sequence select="($style/@font-size,soox:font-from-short-attribute($style)/@font-size,$default-font/@font-size)[1]"/>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Return the font-size using by priority order : 
                <xd:ul>
                    <xd:li>the 'font-weight' attribute value</xd:li>
                    <xd:li>the weight specified in the 'font' attribute</xd:li>
                    <xd:li>the default font-weight</xd:li>
                </xd:ul></xd:p>
        </xd:desc>
        <xd:param name="style">A SOOX 'style' element</xd:param>
        <xd:return>A string that is the font-weight</xd:return>
    </xd:doc>
    <xsl:function name="soox:get-font-weight" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <xsl:sequence select="($style/@font-weight,soox:font-from-short-attribute($style)/@font-weight,$default-font/@font-weight)[1]"/>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Return the font-style using by priority order : 
                <xd:ul>
                    <xd:li>the 'font-style' attribute value</xd:li>
                    <xd:li>the style specified in the 'font' attribute</xd:li>
                    <xd:li>the default font-style</xd:li>
                </xd:ul></xd:p>
        </xd:desc>
        <xd:param name="style">A SOOX 'style' element</xd:param>
        <xd:return>A string that is the font-style</xd:return>
    </xd:doc>
    <xsl:function name="soox:get-font-style" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <xsl:if test="not(index-of(('normal','italic',''),$style/@font-style||''))">
            <xsl:message expand-text="true">Invalid font style specification &quot;{$style/@font-style}&quot;</xsl:message>
        </xsl:if>
        <xsl:sequence select="($style/@font-style,soox:font-from-short-attribute($style)/@font-style,$default-font/@font-style)[1]"/>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Generate the signature string for the cell font style</xd:p>
        </xd:desc>
        <xd:param name="style">the cell style</xd:param>
        <xd:return>a string that uniquely identifies the font style</xd:return>
    </xd:doc>
    <xsl:function name="soox:font-signature" as="xs:string">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <xsl:sequence select="if ($style) then 
                map{
                'family':soox:get-font-family($style),
                'size':soox:get-font-size($style),
                'weight':soox:get-font-weight($style),
                'style':soox:get-font-style($style)
                }=>serialize(map{'method':'adaptive'}) 
            else $default-font-signature"/>
    </xsl:function>
    
    
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
            <xsl:for-each-group select="($default-font,$styles)" group-by="soox:font-signature(.)">
                <xsl:map-entry key="current-grouping-key()">
                    <xsl:variable name="font-style" select="current-group()[1]"/>
                    <sml:font>
                        <sml:sz val="{soox:get-font-size($font-style)}"/>
                        <!--sml:color rgb="FF000000"/-->
                        <sml:name val="{soox:get-font-family($font-style)}"/>
                        <!--sml:family val="2"/>
                        <sml:charset val="238"/-->
                        <xsl:if test="soox:get-font-weight($font-style) = 'bold'">
                            <sml:b/>
                        </xsl:if>
                        <xsl:choose>
                            <xsl:when test="soox:get-font-style($font-style) = 'italic'">
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
            <xd:p>Returns the OOXML fill style index from a font style signature</xd:p>
            <xd:p>The returned index is the position of the signature in the keys of the font styles map</xd:p>
        </xd:desc>
        <xd:param name="styles">A font styles map that associate a signature to a OOXML font element</xd:param>
        <xd:param name="signature">The font style signature to find</xd:param>
        <xd:return>The index of the matching signature or 0(=default font) if not found</xd:return>
    </xd:doc>
    <xsl:function name="soox:font-index-of" as="xs:integer">
        <xsl:param name="styles" as="map(xs:string,element(sml:font))"/>
        <xsl:param name="signature" as="xs:string"/>
        
        <xsl:variable name="matching" as="xs:integer*"
            select="(map:keys($styles)[. ne $default-font-signature])=>index-of($signature)"/>
        <xsl:sequence select="if($matching) then $matching[1] else 0"/>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Generates a table for all unique font styles</xd:p>
        </xd:desc>
        <xd:param name="cellStyles">a sequence of all cell styles</xd:param>
        <xd:return>a font table inside a fonts element</xd:return>
    </xd:doc>
    <xsl:function name="soox:fonts-table" as="element(sml:fonts)">
        <xsl:param name="cellStyles" as="element(s:style)*"/>
        
        <!-- Generates a map {"font-signature": sml:font element} -->
        <xsl:variable name="fontsTablemap" as="map(xs:string,element(sml:font))"
            select="soox:buildFontStyleMap($cellStyles)"/>
        
        <!-- Generates a fonts element containing font elements; first element should be the default one -->
        <sml:fonts count="{count(map:keys($fontsTablemap))}">
            
            <xsl:sequence select="$fontsTablemap($default-font-signature)"/>
            <xsl:for-each select="map:keys($fontsTablemap)[. ne $default-font-signature]">
                <xsl:sequence select="$fontsTablemap(current())"/>
            </xsl:for-each>
        </sml:fonts>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>The default font specification</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:variable name="default-font" as="element(s:style)">
        <s:style font-family="Arial" font-size="12" font-weight="normal" font-style="normal"/>
    </xsl:variable>
    
    <xd:doc>
        <xd:desc>
            <xd:p>The default font signature</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:variable name="default-font-signature" as="xs:string"
        select="$default-font=>soox:font-signature()"/>
    
</xsl:stylesheet>