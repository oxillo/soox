<xsl:stylesheet
  version="3.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
  xmlns:soox="simplify-office-open-xml"
  xmlns:sooxns="simplify-office-open-xml/namespaces"
  xmlns:sml="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:s="soox"
  expand-text="yes"
  exclude-result-prefixes="#all">
  
  <xd:doc scope="stylesheet">
    <xd:desc>
      <xd:p><xd:b>Created on:</xd:b> Nov 13, 2021</xd:p>
      <xd:p><xd:b>Author:</xd:b> Olivier XILLO</xd:p>
      <xd:p/>
    </xd:desc>
  </xd:doc>
  
  <xsl:include href="utils_fonts.xsl"/>
  <xsl:include href="utils_fills.xsl"/>
  <xsl:include href="utils_borders.xsl"/>
  <xsl:include href="utils_numericFormats.xsl"/>
  
  <xsl:mode name="soox:spreadsheet-styles-cascade" on-no-match="shallow-copy"/>
  
  
  <xd:doc>
    <xd:desc>
      <xd:p></xd:p>
    </xd:desc>
    <xd:param name="style"></xd:param>
  </xd:doc>
  <xsl:template match="s:workbook|s:worksheet|s:cell" mode="soox:spreadsheet-styles-cascade">
    <xsl:param name="style" tunnel="yes"/>
    
    <xsl:copy>
      <xsl:copy-of select="@*"/>
      <xsl:apply-templates mode="#current" select="element()">
        <xsl:with-param name="style" select="soox:cascade-style($style,s:style)" tunnel="yes"/>
      </xsl:apply-templates>
    </xsl:copy>
  </xsl:template>
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Rewrite s:style element to remove style attributes that should have been cascaded (workbook and worksheet)
        and to output final style sttributes (after style cascade) for cells</xd:p>
    </xd:desc>
    <xd:param name="style">A map holding the cell style after style cascading (workbook->worksheet->cell) </xd:param>
  </xd:doc>
  <xsl:template match="s:style" mode="soox:spreadsheet-styles-cascade">
    <xsl:param name="style" as="map(*)" tunnel="yes"/>
    
    <xsl:copy>
      <xsl:if test="parent::s:cell">
        <!-- border style -->
        <xsl:attribute name="border-left-style" select="$style('border-left-style')"/>  
        <xsl:attribute name="border-right-style" select="$style('border-right-style')"/>
        <xsl:attribute name="border-top-style" select="$style('border-top-style')"/>
        <xsl:attribute name="border-bottom-style" select="$style('border-bottom-style')"/>
        <!-- border color -->
        <xsl:attribute name="border-left-color" select="$style('border-left-color')"/>  
        <xsl:attribute name="border-right-color" select="$style('border-right-color')"/>
        <xsl:attribute name="border-top-color" select="$style('border-top-color')"/>
        <xsl:attribute name="border-bottom-color" select="$style('border-bottom-color')"/>
        <xsl:variable name="border-signature" select="(
          $style('border-left-style'),$style('border-left-color'),
          $style('border-right-style'),$style('border-right-color'),
          $style('border-top-style'),$style('border-top-color'),
          $style('border-bottom-style'),$style('border-bottom-color'))=>string-join('#')"/>
        <xsl:attribute name="border-signature" select="$border-signature"/>
        <!-- font -->
        <xsl:attribute name="font-family" select="$style('font-family')"/>
        <xsl:attribute name="font-size" select="$style('font-size')"/>
        <xsl:attribute name="font-weight" select="$style('font-weight')"/>
        <xsl:attribute name="font-style" select="$style('font-style')"/>
        <xsl:variable name="font-signature" select="($style('font-family'),$style('font-size'),$style('font-weight'),$style('font-style'))=>string-join('#')"/>
        <xsl:attribute name="font-signature" select="$font-signature"/> 
        <!-- fill -->
        <xsl:attribute name="fill-style" select="$style('fill-style')"/>
        <xsl:attribute name="fill-color" select="$style('fill-color')"/>
        <xsl:variable name="fill-signature" select="($style('fill-style'),$style('fill-color'))=>string-join('#')"/>
        <xsl:attribute name="fill-signature" select="$fill-signature"/>
        <!-- numeric format -->
        <xsl:attribute name="numeric-format" select="$style('numeric-format')"/>
        <xsl:variable name="numeric-format-signature" select="($style('numeric-format'))=>string-join('#')"/>
        <xsl:attribute name="numeric-format-signature" select="$numeric-format-signature"/>
        <!-- style signature -->
        <xsl:attribute name="style-signature" select="($border-signature,$font-signature,$fill-signature,$numeric-format-signature)=>string-join(';')"/>
      </xsl:if>
      <xsl:apply-templates select="element()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  
  <xd:doc>
    <xd:desc>
      <xd:p></xd:p>
    </xd:desc>
    <xd:param name="inherited"></xd:param>
    <xd:param name="local"></xd:param>
  </xd:doc>
  <xsl:function name="soox:cascade-style">
    <xsl:param name="inherited" as="map(*)"/>
    <xsl:param name="local" as="element(s:style)*"/>
      
    <xsl:if test="$local">
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
            $local/@border-left-color,$local/@border-right-color,$local/@border-top-color,$local/@border-bottom-color,
            $local/@font-family,$local/@font-size,$local/@font-style,$local/@font-weight)"/>
        </xsl:map>
      </xsl:variable>
      <xsl:sequence select="map:merge(($inherited=>soox:cascade-fill-style($local),$from-border,$from-border-stylecolor,$from-border-detailed),map{'duplicates':'use-last'})"/>
    </xsl:if>
    <xsl:if test="not($local)">
      <xsl:sequence select="$inherited"/>
    </xsl:if>
  </xsl:function>
  
  
  <xd:doc>
    <xd:desc>
      <xd:p></xd:p>
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
      <xd:p></xd:p>
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
      <xd:p></xd:p>
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
      <xd:p>Generates the 'styles.xml' file information.</xd:p>
    </xd:desc>
    <xd:param name="simple_workbook"></xd:param>
  </xd:doc>
  <xsl:function name="soox:styles.xml" visibility="public" as="map(*)">
    <xsl:param name="simple_workbook" as="element(s:workbook)"/>
      
    <xsl:variable name="content">
      <xsl:variable name="cellStyles" select="$simple_workbook//s:cell/s:style" as="element(s:style)*"/>
      
      <!-- Generates a map {"font-signature": sml:font element} -->
      <xsl:variable name="fontsTablemap" as="map(xs:string,element(sml:font))"
        select="soox:buildFontStyleMap($cellStyles)"/>
      
      <!-- Generates a map {"fill-signature": sml:fill element} -->
      <xsl:variable name="fillsTablemap" as="map(xs:string,element(sml:fill))"
        select="soox:buildFillStyleMap($cellStyles)"/>
      
      <!-- Generates a map {"border-signature": sml:border element} --> 
      <xsl:variable name="bordersTablemap" as="map(xs:string,element(sml:border))"
        select="soox:buildBorderStyleMap($cellStyles)"/>
      
      <!-- Generates a map {"numeric-format-signature": sml:numFmt element} -->
      <xsl:variable name="numericFormatTableMap" as="map(xs:string,xs:integer)"
        select="soox:buildNumFmtMap($cellStyles)"/>
      
      <sml:styleSheet>
        
        <!-- Generate the numeric format table -->
        <xsl:sequence select="soox:numeric-formats-table($numericFormatTableMap)"/>
        
        <!-- Generate the fonts table -->
        <xsl:sequence select="soox:fonts-table($fontsTablemap)"/>
        
        <!-- Generate the fills table -->
        <xsl:sequence select="soox:fills-table($fillsTablemap)"/>
        
        <!-- Generate the borders table -->
        <xsl:sequence select="soox:borders-table($bordersTablemap)"/>
        
        <!-- Generate the cellStyleXfs table -->
        <!-- cellStyleXfs -->
        
        <!-- Generate the cellXfs table -->
        <!-- cellXfs -->
        
        <!-- Generate the cellStyles table -->
        <!-- cellStyles -->
        
        <!-- Generate the dxfs table -->
        <!-- dxfs -->
        
        <!-- Generate the tableStyles table -->
        <!-- tableStyles -->
        
        <!-- Generate the colors table -->
        <!-- colors -->
        
        <!-- Generate the extLst table -->
        <!-- extLst -->
        
        <!--sml:cellStyleXfs count="20">
          <sml:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="true"
            applyAlignment="true" applyProtection="true">
            <sml:alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false"
              indent="0" shrinkToFit="false"/>
            <sml:protection locked="true" hidden="false"/>
          </sml:xf>
          <sml:xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="43" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="41" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="44" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="42" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <sml:xf numFmtId="9" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
        </sml:cellStyleXfs-->
        
        <xsl:variable name="cellStylesTable" as="element(sml:xf)*"
          select="soox:computeStyleCellXfsTable($cellStyles,$fontsTablemap,$fillsTablemap,$bordersTablemap,$numericFormatTableMap)"/>
        <sml:cellXfs count="{1 + $cellStylesTable=>count()}">
          <sml:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="false"
            applyBorder="false" applyAlignment="false" applyProtection="false">
            <sml:alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false"
              indent="0" shrinkToFit="false"/>
            <sml:protection locked="true" hidden="false"/>
          </sml:xf>
          <xsl:sequence select="$cellStylesTable"/> 
        </sml:cellXfs>
        
        <!--sml:cellStyles xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6">
          <sml:cellStyle name="Normal" xfId="0" builtinId="0"/>
          <sml:cellStyle name="Comma" xfId="15" builtinId="3"/>
          <sml:cellStyle name="Comma [0]" xfId="16" builtinId="6"/>
          <sml:cellStyle name="Currency" xfId="17" builtinId="4"/>
          <sml:cellStyle name="Currency [0]" xfId="18" builtinId="7"/>
          <sml:cellStyle name="Percent" xfId="19" builtinId="5"/>
        </sml:cellStyles-->
        
        
        
        
        
        
        
        
        
        <!--cellStyleXfs count="1">
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        </cellStyleXfs>
        <cellXfs count="1">
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        </cellXfs>
        <cellStyles count="1">
          <cellStyle name="Normal" xfId="0" builtinId="0"/>
        </cellStyles>
        <dxfs count="0"/>
        <tableStyles count="0" defaultTableStyle="TableStyleMedium2"
          defaultPivotStyle="PivotStyleLight16"/>
        <extLst>
          <ext uri="{{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}}"
            xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
          </ext>
          <ext uri="{{9260A510-F301-46a8-8635-F512D64BE5F5}}"
            xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
            <x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/>
          </ext>
        </extLst-->
        
      </sml:styleSheet>
    </xsl:variable>
    <xsl:sequence select="map{
      'content': $content,
      'content-type':'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
      'relationship-type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
      }"/>
  </xsl:function>
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Generate a signature that uniquely identifies this style.</xd:p>
      <xd:p>If 2 style element have the same signature, they are considered to define the same style even if their XML content is different</xd:p>
    </xd:desc>
    <xd:param name="style">a style element</xd:param>
    <xd:return>a signature string</xd:return>
  </xd:doc>
  <xsl:function name="soox:styleSignature" as="xs:string">
    <xsl:param name="style" as="element(s:style)?"/>
    
    <xsl:sequence select="map{
      'font':$style/@font-signature,
      'fill':$style/@fill-signature,
      'border':$style/@border-signature,
      'numeric-format':soox:numeric-format-signature($style)
      }=>serialize(map{'method':'adaptive'})"/>
  </xsl:function>
  
  
  
  
  
  
  
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Retrieve a style element from a style elements sequence based on its signature</xd:p>
      <xd:p>If multiple elements match, the first one is returned; if no element match, the empty sequence is returned</xd:p>
    </xd:desc>
    <xd:param name="styles">a sequence of style elements</xd:param>
    <xd:param name="signature">the signature string of the element to find</xd:param>
    <xd:return>The style element that matches the signature string</xd:return>
  </xd:doc>
  <xsl:function name="soox:getStyleFromSignature" as="element(s:style)?">
    <xsl:param name="styles" as="element(s:style)*"/>
    <xsl:param name="signature" as="xs:string"/>
    
    <xsl:variable name="matchingStyles" select="$styles[soox:styleSignature(.)=$signature]"/>
    <xsl:sequence select="$matchingStyles[1]"/>  
  </xsl:function>
  
  
  
  
  <!-- ============================================================================-->
  <!-- =================      Cell Fill style      ================================-->
  <!-- ============================================================================-->
  
  
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Generate a map between the style signature and its index</xd:p>
    </xd:desc>
    <xd:param name="styles"></xd:param>
    <xd:return></xd:return>
  </xd:doc>
  <xsl:function name="soox:buildCellStylesMap" as="map(xs:string,xs:integer)">
    <xsl:param name="styles" as="element(s:style)*"/>
    
    <xsl:map>
      <xsl:for-each-group select="$styles" group-by="soox:styleSignature(.)">
        <xsl:sort order="ascending"/>
        <xsl:map-entry key="current-grouping-key()" select="position()"/>
      </xsl:for-each-group>
    </xsl:map>
  </xsl:function>
  
 
  
  
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Compute the table of unique cell styles</xd:p>
    </xd:desc>
    <xd:param name="styles">The sequence of styles defined in the workbook</xd:param>
    <xd:return>A sequence of xf elements</xd:return>
  </xd:doc>
  <xsl:function name="soox:computeStyleCellXfsTable" as="element(sml:xf)*">
    <xsl:param name="styles" as="element(s:style)*"/>
    <xsl:param name="fontsTablemap" as="map(xs:string,element(sml:font))"/>
    <xsl:param name="fillsTablemap" as="map(xs:string,element(sml:fill))"/>
    <xsl:param name="bordersTablemap" as="map(xs:string,element(sml:border))"/>
    <xsl:param name="numericFormatTableMap" as="map(xs:string,xs:integer)"/>
  
    <!--xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="false"
      applyBorder="false" applyAlignment="false" applyProtection="false">
      <alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false"
        indent="0" shrinkToFit="false"/>
      <protection locked="true" hidden="false"/>
    </xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="true"
      applyBorder="false" applyAlignment="false" applyProtection="false">
      <alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false"
        indent="0" shrinkToFit="false"/>
      <protection locked="true" hidden="false"/>
    </xf-->
    <xsl:for-each-group select="$styles" group-by="soox:styleSignature(.)">
      <xsl:variable name="style" select="current-group()[1]"/>
      
      <sml:xf>
        <xsl:attribute name="numFmtId" select="$numericFormatTableMap=>soox:numeric-format-index-of(soox:numeric-format-signature($style))"/>
        <xsl:attribute name="fontId" select="$fontsTablemap=>soox:font-index-of($style/@font-signature)"/>
        <xsl:attribute name="fillId" select="$fillsTablemap=>soox:index-of($style/@fill-signature)"/>
        <xsl:attribute name="borderId" select="$bordersTablemap=>soox:border-index-of($style/@border-signature)"/>
        <!--xsl:attribute name="xfId" select="(0)[1]"/-->
        <xsl:attribute name="applyFont" select="'true'"/>
        <xsl:attribute name="applyBorder" select="'true'"/>
        <xsl:attribute name="applyAlignment" select="'false'"/>
        <xsl:attribute name="applyProtection" select="'false'"/>
        <!--alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false"
          indent="0" shrinkToFit="false"/>
        <protection locked="true" hidden="false"/-->
      </sml:xf>
    </xsl:for-each-group>
  </xsl:function>
  
  
  <xsl:variable name="default-cell-style" select="map:merge((
    map {
    'border-left-style'  : 'none',
    'border-right-style' : 'none',
    'border-top-style'   : 'none',
    'border-bottom-style': 'none'
    },map {
    'border-left-color'  : 'black',
    'border-right-color' : 'black',
    'border-top-color'   : 'black',
    'border-bottom-color': 'black'
    },map {
    'font-family' : 'Arial',
    'font-size'   : 12,
    'font-weight' : 'normal',
    'font-style'  : 'normal'
    },map {
    'fill-style' : 'none',
    'fill-color'   : 'black'
    },map {
    'numeric-format' : 'General'
    }))"/>
  
</xsl:stylesheet>