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
  </xd:doc>
  <xsl:template match="text()" mode="soox:spreadsheet-styles-cascade"/>
  
  <xd:doc>
    <xd:desc>
      <xd:p></xd:p>
    </xd:desc>
    <xd:param name="default-style"></xd:param>
  </xd:doc>
  <xsl:template match="s:workbook|s:worksheet" mode="soox:spreadsheet-styles-cascade">
    <xsl:param name="default-style" tunnel="yes"/>
    
    <xsl:variable name="updated-default-style" as="map(*)" select="soox:cascade-style($default-style,s:style)"/>
    
    <xsl:copy>
      <xsl:copy-of select="@*"/>
      <xsl:apply-templates mode="#current">
        <xsl:with-param name="default-style" select="$updated-default-style" tunnel="yes"/>
      </xsl:apply-templates>
    </xsl:copy>
  </xsl:template>
  
  <xd:doc>
    <xd:desc>
      <xd:p></xd:p>
    </xd:desc>
    <xd:param name="default-style"></xd:param>
  </xd:doc>
  <xsl:template match="s:cell" mode="soox:spreadsheet-styles-cascade">
    <xsl:param name="default-style" tunnel="yes"/>
    
    <xsl:variable name="s" as="map(*)" select="soox:cascade-style($default-style,s:style)"/>
    <xsl:copy>
      <xsl:element name="style" namespace="soox">
        <xsl:attribute name="border-left" select="$s('border-left-style')"/>  
        <xsl:attribute name="border-right-style" select="$s('border-right-style')"/>
        <xsl:attribute name="border-top-style" select="$s('border-top-style')"/>
        <xsl:attribute name="border-bottom-style" select="$s('border-bottom-style')"/>
        <xsl:attribute name="border-left-color" select="$s('border-left-color')"/>  
        <xsl:attribute name="border-right-color" select="$s('border-right-color')"/>
        <xsl:attribute name="border-top-color" select="$s('border-top-color')"/>
        <xsl:attribute name="border-bottom-color" select="$s('border-bottom-color')"/>
        <xsl:copy-of select="s:style/*"/>
      </xsl:element>
    </xsl:copy>
  </xsl:template>
  
  <xd:doc>
    <xd:desc>
      <xd:p></xd:p>
    </xd:desc>
  </xd:doc>
  <xsl:template match="s:style" mode="soox:spreadsheet-styles-cascade"/>
  
  
  
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
          $local/@border-left-color,$local/@border-right-color,$local/@border-top-color,$local/@border-bottom-color)"/>
      </xsl:map>
    </xsl:variable>
    <xsl:sequence select="map:merge(($inherited,$from-border,$from-border-stylecolor,$from-border-detailed),map{'duplicates':'use-last'})"/>
    </xsl:if>
    <xsl:if test="not($local)">
      <xsl:sequence select="$inherited"/>
    </xsl:if>
  </xsl:function>
  
  
  <!--xsl:template match="s:style" mode="soox:spreadsheet-styles-cascade">
    <xsl:apply-templates select="@border" mode="#current"/>
    <xsl:apply-templates select="(@border-color,@border-style)" mode="#current"/>
    <xsl:copy-of select="(
      @border-style-left,@border-style-right,@border-style-top,@border-style-bottom,
      @border-color-left,@border-color-right,@border-color-top,@border-color-bottom)"/>
  </xsl:template-->
    
  
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
      <xsl:variable name="simple_workbook-cascaded">
        <xsl:apply-templates select="$simple_workbook" mode="soox:spreadsheet-styles-cascade">
          <xsl:with-param name="default-style" as="map(*)" tunnel="yes" 
            select="$default-cell-style"/>
        </xsl:apply-templates>
      </xsl:variable>
      
      <xsl:variable name="cellStyles" select="$simple_workbook-cascaded//s:cell/s:style" as="element(s:style)*"/>
      <sml:styleSheet>
        
        <!-- Generate the numeric format table -->
        <xsl:sequence select="soox:numeric-formats-table($cellStyles)"/>
        
        <!-- Generate the fonts table -->
        <xsl:sequence select="soox:fonts-table($cellStyles)"/>
        
        <!-- Generate the fills table -->
        <xsl:sequence select="soox:fills-table($cellStyles)"/>
        
        <!-- Generate the borders table -->
        <xsl:sequence select="soox:borders-table($cellStyles)"/>
        
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
          select="soox:computeStyleCellXfsTable($cellStyles)"/>
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
      'font':soox:font-signature($style),
      'fill':soox:fill-signature($style),
      'border':soox:border-signature($style),
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
        <xsl:attribute name="numFmtId" select="soox:buildNumFmtMap($styles)=>soox:numeric-format-index-of(soox:numeric-format-signature($style))"/>
        <xsl:attribute name="fontId" select="soox:buildFontStyleMap($styles)=>soox:font-index-of(soox:font-signature($style))"/>
        <xsl:attribute name="fillId" select="soox:buildFillStyleMap($styles)=>soox:fill-index-of(soox:fill-signature($style))"/>
        <xsl:attribute name="borderId" select="soox:buildBorderStyleMap($styles)=>soox:border-index-of(soox:border-signature($style))"/>
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
    }))"/>
  
</xsl:stylesheet>