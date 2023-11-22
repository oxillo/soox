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
  <xsl:include href="utils_borders.xsl"/>
  
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
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <numFmts xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1">
          <numFmt numFmtId="164" formatCode="General"/>
        </numFmts>
        
        <xsl:variable name="fontsTablemap" as="map(xs:string,element(sml:font))"
          select="soox:buildFontStyleMap($cellStyles)"/>
        <fonts count="{count(map:keys($fontsTablemap))}">
          <xsl:variable name="default-font-signature" select="$default-font=>soox:fontSignature()"/>
          <xsl:sequence select="$fontsTablemap($default-font-signature)"/>
          <xsl:for-each select="map:keys($fontsTablemap)[. ne $default-font-signature]">
            <xsl:sequence select="$fontsTablemap(current())"/>
          </xsl:for-each>
        </fonts>
        
        <xsl:variable name="fillsTablemap" as="map(xs:string,element(sml:fill))"
          select="soox:buildFillStyleMap($cellStyles)"/>
        <fills count="{1 + count(map:keys($fillsTablemap))}">
          <fill>
            <patternFill patternType="none"/>
          </fill>
          
          <xsl:for-each select="map:keys($fillsTablemap)">
            <xsl:sequence select="$fillsTablemap(current())"/>
          </xsl:for-each>
        </fills>
        
        <xsl:variable name="bordersTablemap" as="map(xs:string,element(sml:border))"
          select="soox:buildBorderStyleMap($cellStyles)"/>
        <borders count="{1 + count(map:keys($bordersTablemap))}">
          <xsl:variable name="default-border-signature" select="$default-border=>soox:border-signature()"/>
          <xsl:sequence select="$bordersTablemap($default-border-signature)"/>
          <xsl:for-each select="map:keys($bordersTablemap)[. ne $default-border-signature]">
            <xsl:sequence select="$bordersTablemap(current())"/>
          </xsl:for-each>
        </borders>
        
        <cellStyleXfs xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="20">
          <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="true"
            applyAlignment="true" applyProtection="true">
            <alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false"
              indent="0" shrinkToFit="false"/>
            <protection locked="true" hidden="false"/>
          </xf>
          <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="43" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="41" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="44" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="42" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
          <xf numFmtId="9" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false"
            applyAlignment="false" applyProtection="false"/>
        </cellStyleXfs>
        
        <xsl:variable name="cellStylesTable" as="element(sml:xf)*"
          select="soox:computeStyleCellXfsTable($cellStyles)"/>
        <cellXfs count="{1 + $cellStylesTable=>count()}">
          <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="false"
            applyBorder="false" applyAlignment="false" applyProtection="false">
            <alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false"
              indent="0" shrinkToFit="false"/>
            <protection locked="true" hidden="false"/>
          </xf>
          <xsl:sequence select="$cellStylesTable"/> 
        </cellXfs>
        
        <cellStyles xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6">
          <cellStyle name="Normal" xfId="0" builtinId="0"/>
          <cellStyle name="Comma" xfId="15" builtinId="3"/>
          <cellStyle name="Comma [0]" xfId="16" builtinId="6"/>
          <cellStyle name="Currency" xfId="17" builtinId="4"/>
          <cellStyle name="Currency [0]" xfId="18" builtinId="7"/>
          <cellStyle name="Percent" xfId="19" builtinId="5"/>
        </cellStyles>
        
        
        
        
        
        
        
        
        
        <!--fonts count="1">
          <font>
            <sz val="11"/>
            <color theme="1"/>
            <name val="Arial"/>
            <family val="2"/>
          </font>
        </fonts>
        <cellStyleXfs count="1">
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
        
      </styleSheet>
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
      'font':soox:fontSignature($style),
      'fill':soox:fillSignature($style),
      'border':soox:border-signature($style)
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
      <xd:p>Generate the signature string for the cell fill style</xd:p>
    </xd:desc>
    <xd:param name="style">the cell style</xd:param>
    <xd:return>a string that uniquely identifies the fill style</xd:return>
  </xd:doc>
  <xsl:function name="soox:fillSignature" as="xs:string">
    <xsl:param name="style" as="element(s:style)?"/>
    
    <xsl:sequence select="if ($style/@fill-color) then 'solid('||soox:parseColor($style/@fill-color)||')' else 'none'"/>
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
      <xsl:for-each-group select="$styles" group-by="soox:fillSignature(.)">
        <xsl:sort order="ascending"/>
        
        <xsl:if test="current-grouping-key() ne 'none'">
          <xsl:map-entry key="current-grouping-key()">
            <xsl:variable name="color" select="soox:parseColor(current-group()[1]/@fill-color)"/>
            <fill>
              <patternFill patternType="solid">
                <fgColor rgb="{$color}"/>
                <bgColor rgb="{$color}"/>
              </patternFill>
            </fill>
          </xsl:map-entry>
        </xsl:if>
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
  <xsl:function name="soox:getFillStyleIndex" as="xs:integer">
    <xsl:param name="signature" as="xs:string"/>
    <xsl:param name="styles" as="map(xs:string,element(sml:fill))"/>
      
    <xsl:variable name="matching" as="xs:integer*"
      select="map:keys($styles)=>index-of($signature)"/>
    <xsl:sequence select="if($matching) then $matching[1] else 0"/>
  </xsl:function>
  
  
  
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
      
      <xf>
        <xsl:attribute name="numFmtId" select="(164)[1]"/>
        <xsl:attribute name="fontId" select="soox:buildFontStyleMap($styles)=>soox:font-index-of(soox:fontSignature($style))"/>
        <xsl:attribute name="fillId" select="soox:fillSignature($style)=>soox:getFillStyleIndex(soox:buildFillStyleMap($styles))"/>
        <xsl:attribute name="borderId" select="soox:buildBorderStyleMap($styles)=>soox:border-index-of(soox:border-signature($style))"/>
        <xsl:attribute name="xfId" select="(0)[1]"/>
        <xsl:attribute name="applyFont" select="'true'"/>
        <xsl:attribute name="applyBorder" select="'false'"/>
        <xsl:attribute name="applyAlignment" select="'false'"/>
        <xsl:attribute name="applyProtection" select="'false'"/>
        <!--alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false"
          indent="0" shrinkToFit="false"/>
        <protection locked="true" hidden="false"/-->
      </xf>
    </xsl:for-each-group>
  </xsl:function>
  
  
</xsl:stylesheet>