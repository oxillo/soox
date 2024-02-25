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
            <xd:p>Generates a 'style' element with detailed border attributes from a style with the shortened 'border' attrinbute</xd:p>
            <xd:p><![CDATA[<style font="12 'Arial'"/> generates <style font-family="Arial" font-size="12"/>]]></xd:p>
        </xd:desc>
        <xd:param name="style">A 'style' element</xd:param>
        <xd:return>A 'style' element</xd:return>
    </xd:doc>
    <xsl:function name="soox:expand-border-attributes" as="element(s:style)">
        <xsl:param name="style" as="element(s:style)?"/>
        
        <!--TODO-->
        <xsl:variable name="from-border" as="element(s:style)">
            <xsl:variable name="tokens" select="($style/@border=>normalize-space()=>tokenize(' '))"/>
            <s:style>
                <xsl:if test="$tokens[1]=$borders-styles">
                    <xsl:attribute name="border-left-style" select="$tokens[1]"/>
                    <xsl:attribute name="border-right-style" select="$tokens[1]"/>
                    <xsl:attribute name="border-top-style" select="$tokens[1]"/>
                    <xsl:attribute name="border-bottom-style" select="$tokens[1]"/>
                </xsl:if>
                <xsl:if test="$tokens[2] and not($tokens[2]=>soox:parse-color('invalid') = 'invalid')">
                    <xsl:attribute name="border-left-color" select="$tokens[2]"/>
                    <xsl:attribute name="border-right-color" select="$tokens[2]"/>
                    <xsl:attribute name="border-top-color" select="$tokens[2]"/>
                    <xsl:attribute name="border-bottom-color" select="$tokens[2]"/>
                </xsl:if>
            </s:style>
        </xsl:variable>
        
        <xsl:variable name="from-border-style" as="element(s:style)">
            <xsl:variable name="tokens" select="($style/@border-style=>normalize-space()=>tokenize(' '))"/>
            <s:style>
                <xsl:choose>
                    <xsl:when test="count($tokens)=1">
                        <xsl:if test="$tokens[1]=$borders-styles">
                            <xsl:attribute name="border-left-style" select="$tokens[1]"/>
                            <xsl:attribute name="border-right-style" select="$tokens[1]"/>
                            <xsl:attribute name="border-top-style" select="$tokens[1]"/>
                            <xsl:attribute name="border-bottom-style" select="$tokens[1]"/>
                        </xsl:if>
                    </xsl:when>
                    <xsl:when test="count($tokens)=2">
                        <xsl:if test="$tokens[1]=$borders-styles">
                            <xsl:attribute name="border-top-style" select="$tokens[1]"/>
                            <xsl:attribute name="border-bottom-style" select="$tokens[1]"/>
                        </xsl:if>
                        <xsl:if test="$tokens[2]=$borders-styles">
                            <xsl:attribute name="border-left-style" select="$tokens[2]"/>
                            <xsl:attribute name="border-right-style" select="$tokens[2]"/>
                        </xsl:if>
                    </xsl:when>
                    <xsl:when test="count($tokens)=3">
                        <xsl:if test="$tokens[1]=$borders-styles">
                            <xsl:attribute name="border-top-style" select="$tokens[1]"/>
                        </xsl:if>
                        <xsl:if test="$tokens[2]=$borders-styles">
                            <xsl:attribute name="border-left-style" select="$tokens[2]"/>
                            <xsl:attribute name="border-right-style" select="$tokens[2]"/>
                        </xsl:if>
                        <xsl:if test="$tokens[3]=$borders-styles">
                            <xsl:attribute name="border-bottom-style" select="$tokens[3]"/>
                        </xsl:if>
                    </xsl:when>
                    <xsl:when test="count($tokens)=4">
                        <xsl:if test="$tokens[1]=$borders-styles">
                            <xsl:attribute name="border-top-style" select="$tokens[1]"/>
                        </xsl:if>
                        <xsl:if test="$tokens[2]=$borders-styles">
                            <xsl:attribute name="border-right-style" select="$tokens[2]"/>
                        </xsl:if>
                        <xsl:if test="$tokens[3]=$borders-styles">
                            <xsl:attribute name="border-bottom-style" select="$tokens[3]"/>
                        </xsl:if>
                        <xsl:if test="$tokens[4]=$borders-styles">
                            <xsl:attribute name="border-left-style" select="$tokens[4]"/>
                        </xsl:if>
                    </xsl:when>
                    <xsl:otherwise/>
                </xsl:choose>
            </s:style>
        </xsl:variable>
        
        <xsl:variable name="from-border-color" as="element(s:style)">
            <xsl:variable name="tokens" select="($style/@border-color=>normalize-space()=>tokenize(' '))"/>
            <s:style>
                <xsl:choose>
                    <xsl:when test="count($tokens)=1">
                        <xsl:attribute name="border-left-color" select="$tokens[1]"/>
                        <xsl:attribute name="border-right-color" select="$tokens[1]"/>
                        <xsl:attribute name="border-top-color" select="$tokens[1]"/>
                        <xsl:attribute name="border-bottom-color" select="$tokens[1]"/>
                    </xsl:when>
                    <xsl:when test="count($tokens)=2">
                        <xsl:attribute name="border-top-color" select="$tokens[1]"/>
                        <xsl:attribute name="border-bottom-color" select="$tokens[1]"/>
                        <xsl:attribute name="border-left-color" select="$tokens[2]"/>
                        <xsl:attribute name="border-right-color" select="$tokens[2]"/>
                    </xsl:when>
                    <xsl:when test="count($tokens)=3">
                        <xsl:attribute name="border-top-color" select="$tokens[1]"/>
                        <xsl:attribute name="border-left-color" select="$tokens[2]"/>
                        <xsl:attribute name="border-right-color" select="$tokens[2]"/>
                        <xsl:attribute name="border-bottom-color" select="$tokens[3]"/>
                    </xsl:when>
                    <xsl:when test="count($tokens)=4">
                        <xsl:attribute name="border-top-color" select="$tokens[1]"/>
                        <xsl:attribute name="border-right-color" select="$tokens[2]"/>
                        <xsl:attribute name="border-bottom-color" select="$tokens[3]"/>
                        <xsl:attribute name="border-left-color" select="$tokens[4]"/>
                    </xsl:when>
                    <xsl:otherwise/>
                </xsl:choose>
            </s:style>
        </xsl:variable>
        
        <xsl:variable name="from-border-leftrighttopbottom" as="element(s:style)+">
            <s:style>
                <xsl:variable name="tokens" select="$style/@border-left=>normalize-space()=>tokenize(' ')"/>
                <xsl:if test="$tokens[1]=$borders-styles">
                    <xsl:attribute name="border-left-style" select="$tokens[1]"/>
                </xsl:if>
                <xsl:if test="$tokens[2] and not($tokens[2]=>soox:parse-color('invalid') = 'invalid')">
                    <xsl:attribute name="border-left-color" select="$tokens[2]"/>
                </xsl:if>
            </s:style>
            <s:style>
                <xsl:variable name="tokens" select="$style/@border-right=>normalize-space()=>tokenize(' ')"/>
                <xsl:if test="$tokens[1]=$borders-styles">
                    <xsl:attribute name="border-right-style" select="$tokens[1]"/>
                </xsl:if>
                <xsl:if test="$tokens[2] and not($tokens[2]=>soox:parse-color('invalid') = 'invalid')">
                    <xsl:attribute name="border-right-color" select="$tokens[2]"/>
                </xsl:if>
            </s:style>
            <s:style>
                <xsl:variable name="tokens" select="$style/@border-top=>normalize-space()=>tokenize(' ')"/>
                <xsl:if test="$tokens[1]=$borders-styles">
                    <xsl:attribute name="border-top-style" select="$tokens[1]"/>
                </xsl:if>
                <xsl:if test="$tokens[2] and not($tokens[2]=>soox:parse-color('invalid') = 'invalid')">
                    <xsl:attribute name="border-top-color" select="$tokens[2]"/>
                </xsl:if>
            </s:style>
            <s:style>
                <xsl:variable name="tokens" select="$style/@border-bottom=>normalize-space()=>tokenize(' ')"/>
                <xsl:if test="$tokens[1]=$borders-styles">
                    <xsl:attribute name="border-bottom-style" select="$tokens[1]"/>
                </xsl:if>
                <xsl:if test="$tokens[2] and not($tokens[2]=>soox:parse-color('invalid') = 'invalid')">
                    <xsl:attribute name="border-bottom-color" select="$tokens[2]"/>
                </xsl:if>
            </s:style>
        </xsl:variable>
        
        <xsl:variable name="from-border-detailed" as="element(s:style)">
            <s:style>
                <xsl:copy-of select="$style/attribute::*[name()= $borders]"/>
                <!--xsl:for-each select="$borders">
                    <xsl:copy-of select="$style/attribute::*[name()= current()]"/>
                </xsl:for-each-->
                
                <!--xsl:for-each select="('left','right','top','bottom')!concat('border-',.)">
                    <xsl:for-each select="('style','color')!concat(current(),'-',.)">
                        <xsl:copy-of select="$style/attribute::*[name()= current()]"/>
                    </xsl:for-each>
                </xsl:for-each-->
            </s:style>
        </xsl:variable>
        <xsl:sequence select="($from-border-detailed,$from-border-leftrighttopbottom,$from-border-color,$from-border-style,$from-border)=>soox:cascade-border-style()"/>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Generate a single 'style' element with cascaded border styles</xd:p>
            <xd:p>It goes over all border attributes and retain the first defined one from the styles sequence</xd:p>
        </xd:desc>
        <xd:param name="styles">A sequence of style elements. The first one is the most important</xd:param>
        <xd:return>A style element with effective border attributes</xd:return>
    </xd:doc>
    <xsl:function name="soox:cascade-border-style" as="element(s:style)">
        <xsl:param name="styles" as="element(s:style)+"/>
        
        <s:style>
            <xsl:attribute name="border-top-style" select="($styles/@border-top-style,$default-border/@border-top-style)[1]"/>
            <xsl:attribute name="border-left-style" select="($styles/@border-left-style,$default-border/@border-left-style)[1]"/>
            <xsl:attribute name="border-right-style" select="($styles/@border-right-style,$default-border/@border-right-style)[1]"/>
            <xsl:attribute name="border-bottom-style" select="($styles/@border-bottom-style,$default-border/@border-bottom-style)[1]"/>
            <xsl:attribute name="border-top-color" select="($styles/@border-top-color,$default-border/@border-top-color)[1]"/>
            <xsl:attribute name="border-left-color" select="($styles/@border-left-color,$default-border/@border-left-color)[1]"/>
            <xsl:attribute name="border-right-color" select="($styles/@border-right-color,$default-border/@border-right-color)[1]"/>
            <xsl:attribute name="border-bottom-color" select="($styles/@border-bottom-color,$default-border/@border-bottom-color)[1]"/>
        </s:style>
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
        <xsl:param name="expanded-style" as="element(s:style)?"/>
        
        <xsl:variable name="border-style" as="xs:string"
            select="($style/@border-left-style,$expanded-style/@border-left-style,'none')[1]"/>
        <xsl:choose>
            <xsl:when test="$border-style='none'">none</xsl:when>
            <xsl:otherwise>
                <xsl:variable name="border-color" as="xs:string"
                    select="($style/@border-left-color,$expanded-style/@border-left-color,'black')[1]=>soox:parse-color()"/>
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
        <xsl:param name="expanded-style" as="element(s:style)?"/>
        
        
        <xsl:variable name="border-style" as="xs:string"
            select="($style/@border-right-style,$expanded-style/@border-right-style,'none')[1]"/>
        <xsl:choose>
            <xsl:when test="$border-style='none'">none</xsl:when>
            <xsl:otherwise>
                <xsl:variable name="border-color" as="xs:string"
                    select="($style/@border-right-color,$expanded-style/@border-right-color,'black')[1]=>soox:parse-color()"/>
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
        <xsl:param name="expanded-style" as="element(s:style)?"/>
        
        <xsl:variable name="border-style" as="xs:string"
            select="($style/@border-top-style,$expanded-style/@border-top-style,'none')[1]"/>
        <xsl:choose>
            <xsl:when test="$border-style='none'">none</xsl:when>
            <xsl:otherwise>
                <xsl:variable name="border-color" as="xs:string"
                    select="($style/@border-top-color,$expanded-style/@border-top-color,'black')[1]=>soox:parse-color()"/>
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
        <xsl:param name="expanded-style" as="element(s:style)?"/>
        
        <xsl:variable name="border-style" as="xs:string"
            select="($style/@border-bottom-style,$expanded-style/@border-bottom-style,'none')[1]"/>
        <xsl:choose>
            <xsl:when test="$border-style='none'">none</xsl:when>
            <xsl:otherwise>
                <xsl:variable name="border-color" as="xs:string"
                    select="($style/@border-bottom-color,$expanded-style/@border-bottom-color,'black')[1]=>soox:parse-color()"/>
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
        
        <xsl:variable name="expanded-style" select="soox:expand-border-attributes($style)"/>
        <xsl:sequence select="if ($style) then 
                map{
                'left':soox:border-left($style,$expanded-style),
                'right':soox:border-right($style,$expanded-style),
                'top':soox:border-top($style,$expanded-style),
                'bottom':soox:border-bottom($style,$expanded-style)
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
                    <xsl:variable name="expanded-style" select="soox:expand-border-attributes($border-style)"/>
                    <sml:border>
                        <sml:left>
                            <xsl:variable name="args" select="soox:border-left($border-style,$expanded-style)=>tokenize('#')"/>
                            <xsl:if test="$args[1] ne 'none'">
                                <xsl:attribute name="style" select="$args[1]"/>
                                <sml:color rgb="{$args[2]}"/>
                            </xsl:if>
                        </sml:left>
                        <sml:right>
                            <xsl:variable name="args" select="soox:border-right($border-style,$expanded-style)=>tokenize('#')"/>
                            <xsl:if test="$args[1] ne 'none'">
                                <xsl:attribute name="style" select="$args[1]"/>
                                <sml:color rgb="{$args[2]}"/>
                            </xsl:if>
                        </sml:right>
                        <sml:top>
                            <xsl:variable name="args" select="soox:border-top($border-style,$expanded-style)=>tokenize('#')"/>
                            <xsl:if test="$args[1] ne 'none'">
                                <xsl:attribute name="style" select="$args[1]"/>
                                <sml:color rgb="{$args[2]}"/>
                            </xsl:if>
                        </sml:top>
                        <sml:bottom>
                            <xsl:variable name="args" select="soox:border-bottom($border-style,$expanded-style)=>tokenize('#')"/>
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
            <xd:p>Generates a table for all unique border styles</xd:p>
        </xd:desc>
        <xd:param name="cellStyles">a sequence of all the cell styles</xd:param>
        <xd:return>a border table inside a borders element</xd:return>
    </xd:doc>
    <xsl:function name="soox:borders-table" as="element(sml:borders)">
        <xsl:param name="cellStyles" as="element(s:style)*"/>
        
        <!-- Generates a map {"border-signature": sml:border element} --> 
        <xsl:variable name="bordersTablemap" as="map(xs:string,element(sml:border))"
            select="soox:buildBorderStyleMap($cellStyles)"/>
        
        <!-- Generates a borders element containing border elements; first element should be the default one -->
        <sml:borders count="{1 + count(map:keys($bordersTablemap))}">
            
            <xsl:variable name="default-border-signature" select="$default-border=>soox:border-signature()"/>
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
            border-left-color="black" border-right-color="black" border-top-color="black" border-bottom-color="black"/>
    </xsl:variable>
    
    <xsl:variable name="borders-styles" as="xs:string+"
        select="('inherit','none','thin','medium','dashed','dotted','thick','double','hair',
        'mediumDashed','dashDot','mediumDashDot','dashDotDot','mediumDashDotDot','slantDashDot')"/>
    
    <xsl:variable name="borders" select="('border-left-style','border-left-color',
        'border-right-style','border-right-color',
        'border-top-style','border-top-color',
        'border-bottom-style','border-bottom-color')"/>
    
</xsl:stylesheet>