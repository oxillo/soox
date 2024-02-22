<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:map="http://www.w3.org/2005/xpath-functions/map"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
    xmlns:soox="simplify-office-open-xml"
    extension-element-prefixes="soox"
    exclude-result-prefixes="#all"
    expand-text="yes"
    version="3.0">
    <xd:doc scope="stylesheet">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> Nov 12, 2023</xd:p>
            <xd:p><xd:b>Author:</xd:b> Olivier</xd:p>
            <xd:p>Helper functions for color manipulation</xd:p>
        </xd:desc>
    </xd:doc>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Return a unique ARGB color representation from the various color specification supported by soox</xd:p>
            <xd:p>It currently supports : 
                <xd:ul>
                    <xd:li>hex RGB : #RRGGBB</xd:li>
                    <xd:li>hex RGBA : #RRGGBBAA</xd:li>
                    <xd:li>hex RGB (shortened): #RGB</xd:li>
                    <xd:li>hex RGBA (shortened): #RGBA</xd:li>
                    <xd:li>CSS color names : 'red', 'aqua',...</xd:li>
                </xd:ul>
            </xd:p>
            <xd:p>If the string cannot be parsed, the function returns the default color</xd:p>
        </xd:desc>
        <xd:param name="str">a string</xd:param>
        <xd:param name="default">a string representing a ARGB color that is used in case of parsing error.</xd:param>
        <xd:return>a string that is the hexadecimal ARGB value.(eg. FF123456)</xd:return>
    </xd:doc>
    <xsl:function name="soox:parse-color" as="xs:string">
        <xsl:param name="str" as="xs:string"/>
        <xsl:param name="default" as="xs:string"/>
        
        <!--Remove white space from the string -->
        <xsl:variable name="color" as="xs:string" select="replace($str,'\s','')"/>
        <!-- Try to determine the color specification format -->
        <xsl:choose>
            <!-- Hex RGB -->
            <xsl:when test="matches($color,'^#([A-Fa-f0-9]{6})$')">
                <xsl:variable name="r" select="substring($color,2,2)=>upper-case()"/>
                <xsl:variable name="g" select="substring($color,4,2)=>upper-case()"/>
                <xsl:variable name="b" select="substring($color,6,2)=>upper-case()"/>
                <xsl:sequence select="concat('FF',$r,$g,$b)"/>    
            </xsl:when>
            <!-- Hex RGB (shortened)-->
            <xsl:when test="matches($color,'^#([A-Fa-f0-9]{3})$')">
                <xsl:variable name="r" select="substring($color,2,1)=>upper-case()"/>
                <xsl:variable name="g" select="substring($color,3,1)=>upper-case()"/>
                <xsl:variable name="b" select="substring($color,4,1)=>upper-case()"/>
                <xsl:sequence select="concat('FF',$r,$r,$g,$g,$b,$b)"/>    
            </xsl:when>
            <!-- Hex RGBA -->
            <xsl:when test="matches($color,'^#([A-Fa-f0-9]{8})$')">
                <xsl:variable name="r" select="substring($color,2,2)=>upper-case()"/>
                <xsl:variable name="g" select="substring($color,4,2)=>upper-case()"/>
                <xsl:variable name="b" select="substring($color,6,2)=>upper-case()"/>
                <xsl:variable name="a" select="substring($color,8,2)=>upper-case()"/>
                <xsl:sequence select="concat($a,$r,$g,$b)"/>    
            </xsl:when>
            <!-- Hex RGBA (shortened)-->
            <xsl:when test="matches($color,'^#([A-Fa-f0-9]{4})$')">
                <xsl:variable name="r" select="substring($color,2,1)=>upper-case()"/>
                <xsl:variable name="g" select="substring($color,3,1)=>upper-case()"/>
                <xsl:variable name="b" select="substring($color,4,1)=>upper-case()"/>
                <xsl:variable name="a" select="substring($color,5,1)=>upper-case()"/>
                <xsl:sequence select="concat($a,$a,$r,$r,$g,$g,$b,$b)"/>    
            </xsl:when>
            <!-- CSS color name--> 
            <xsl:when test="$color=map:keys($CssColors)">
                <xsl:sequence select="$CssColors($color)"/>    
            </xsl:when>
            <!-- rgb(r,g,b) specification -->
            <!-- TODO : need to convert to 255 to FF -->
            <!--xsl:when test="matches($color,'^rgb\(\d+,\d+,\d+\)$')">
        <xsl:variable name="tokens" select="$color=>substring-after('rgb(')=>substring-before(')')=>tokenize(',')"/>
        <xsl:variable name="r" select="$tokens[1]"/> 
        <xsl:variable name="g" select="$tokens[2]"/>
        <xsl:variable name="b" select="$tokens[3]"/>
        <xsl:sequence select="concat('FF',$r,$g,$b)"/>    
      </xsl:when-->
            <!-- Unknow format -->
            <xsl:otherwise>
                <xsl:message>Color specification "<xsl:value-of select="$color"/>" cannot be parsed.</xsl:message>
                <xsl:sequence select="$default"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Return a unique ARGB color representation from the various color specification supported by soox</xd:p>
            <xd:p>It currently supports : 
                <xd:ul>
                    <xd:li>hex RGB : #RRGGBB</xd:li>
                    <xd:li>hex RGBA : #RRGGBBAA</xd:li>
                    <xd:li>hex RGB (shortened): #RGB</xd:li>
                    <xd:li>hex RGBA (shortened): #RGBA</xd:li>
                    <xd:li>CSS color names : 'red', 'aqua',...</xd:li>
                </xd:ul>
            </xd:p>
            <xd:p>If the string cannot be parsed, the function returns the color 'black' color</xd:p>
        </xd:desc>
        <xd:param name="str">a string</xd:param>
        <xd:return>a string that is the hexadecimal ARGB value.(eg. FF123456)</xd:return>
    </xd:doc>
    <xsl:function name="soox:parse-color" as="xs:string">
        <xsl:param name="str" as="xs:string"/>
        
        <xsl:sequence select="soox:parse-color($str,'FF000000')"/>
    </xsl:function>
    
    <xd:doc>
        <xd:desc>
            <xd:p>CSS color definitions</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:variable name="CssColors" as="map(xs:string,xs:string)"
        select="map{
        'aliceblue':'FFF0F8FF',
        'antiquewhite':'FFFAEBD7',
        'aqua':'FF00FFFF',
        'aquamarine':'FF7FFFD4',
        'azure':'FFF0FFFF',
        'beige':'FFF5F5DC',
        'bisque':'FFFFE4C4',
        'black':'FF000000',
        'blanchedalmond':'FFFFEBCD',
        'blue':'FF0000FF',
        'blueviolet':'FF8A2BE2',
        'brown':'FFA52A2A',
        'burlywood':'FFDEB887',
        'cadetblue':'FF5F9EA0',
        'chartreuse':'FF7FFF00',
        'chocolate':'FFD2691E',
        'coral':'FFFF7F50',
        'cornflowerblue':'FF6495ED',
        'cornsilk':'FFFFF8DC',
        'crimson':'FFDC143C',
        'cyan':'FF00FFFF',
        'darkblue':'FF00008B',
        'darkcyan':'FF008B8B',
        'darkgoldenrod':'FFB8860B',
        'darkgray':'FFA9A9A9',
        'darkgreen':'FF006400',
        'darkkhaki':'FFBDB76B',
        'darkmagenta':'FF8B008B',
        'darkolivegreen':'FF556B2F',
        'darkorange':'FFFF8C00',
        'darkorchid':'FF9932CC',
        'darkred':'FF8B0000',
        'darksalmon':'FFE9967A',
        'darkseagreen':'FF8FBC8F',
        'darkslateblue':'FF483D8B',
        'darkslategray':'FF2F4F4F',
        'darkturquoise':'FF00CED1',
        'darkviolet':'FF9400D3',
        'deeppink':'FFFF1493',
        'deepskyblue':'FF00BFFF',
        'dimgray':'FF696969',
        'dodgerblue':'FF1E90FF',
        'firebrick':'FFB22222',
        'floralwhite':'FFFFFAF0',
        'forestgreen':'FF228B22',
        'fuchsia':'FFFF00FF',
        'gainsboro':'FFDCDCDC',
        'ghostwhite':'FFF8F8FF',
        'gold':'FFFFD700',
        'goldenrod':'FFDAA520',
        'gray':'FF7F7F7F',
        'green':'FF008000',
        'greenyellow':'FFADFF2F',
        'honeydew':'FFF0FFF0',
        'hotpink':'FFFF69B4',
        'indianred':'FFCD5C5C',
        'indigo':'FF4B0082',
        'ivory':'FFFFFFF0',
        'khaki':'FFF0E68C',
        'lavender':'FFE6E6FA',
        'lavenderblush':'FFFFF0F5',
        'lawngreen':'FF7CFC00',
        'lemonchiffon':'FFFFFACD',
        'lightblue':'FFADD8E6',
        'lightcoral':'FFF08080',
        'lightcyan':'FFE0FFFF',
        'lightgoldenrodyellow':'FFFAFAD2',
        'lightgreen':'FF90EE90',
        'lightgrey':'FFD3D3D3',
        'lightpink':'FFFFB6C1',
        'lightsalmon':'FFFFA07A',
        'lightseagreen':'FF20B2AA',
        'lightskyblue':'FF87CEFA',
        'lightslategray':'FF778899',
        'lightsteelblue':'FFB0C4DE',
        'lightyellow':'FFFFFFE0',
        'lime':'FF00FF00',
        'limegreen':'FF32CD32',
        'linen':'FFFAF0E6',
        'magenta':'FFFF00FF',
        'maroon':'FF800000',
        'mediumaquamarine':'FF66CDAA',
        'mediumblue':'FF0000CD',
        'mediumorchid':'FFBA55D3',
        'mediumpurple':'FF9370DB',
        'mediumseagreen':'FF3CB371',
        'mediumslateblue':'FF7B68EE',
        'mediumspringgreen':'FF00FA9A',
        'mediumturquoise':'FF48D1CC',
        'mediumvioletred':'FFC71585',
        'midnightblue':'FF191970',
        'mintcream':'FFF5FFFA',
        'mistyrose':'FFFFE4E1',
        'moccasin':'FFFFE4B5',
        'navajowhite':'FFFFDEAD',
        'navy':'FF000080',
        'navyblue':'FF9FAFDF',
        'oldlace':'FFFDF5E6',
        'olive':'FF808000',
        'olivedrab':'FF6B8E23',
        'orange':'FFFFA500',
        'orangered':'FFFF4500',
        'orchid':'FFDA70D6',
        'palegoldenrod':'FFEEE8AA',
        'palegreen':'FF98FB98',
        'paleturquoise':'FFAFEEEE',
        'palevioletred':'FFDB7093',
        'papayawhip':'FFFFEFD5',
        'peachpuff':'FFFFDAB9',
        'peru':'FFCD853F',
        'pink':'FFFFC0CB',
        'plum':'FFDDA0DD',
        'powderblue':'FFB0E0E6',
        'purple':'FF800080',
        'red':'FFFF0000',
        'rosybrown':'FFBC8F8F',
        'royalblue':'FF4169E1',
        'saddlebrown':'FF8B4513',
        'salmon':'FFFA8072',
        'sandybrown':'FFFA8072',
        'seagreen':'FF2E8B57',
        'seashell':'FFFFF5EE',
        'sienna':'FFA0522D',
        'silver':'FFC0C0C0',
        'skyblue':'FF87CEEB',
        'slateblue':'FF6A5ACD',
        'slategray':'FF708090',
        'snow':'FFFFFAFA',
        'springgreen':'FF00FF7F',
        'steelblue':'FF4682B4',
        'tan':'FFD2B48C',
        'teal':'FF008080',
        'thistle':'FFD8BFD8',
        'tomato':'FFFF6347',
        'turquoise':'FF40E0D0',
        'violet':'FFEE82EE',
        'wheat':'FFF5DEB3',
        'white':'FFFFFFFF',
        'whitesmoke':'FFF5F5F5',
        'yellow':'FFFFFF00',
        'yellowgreen':'FF9ACD32'
        }"/>
</xsl:stylesheet>