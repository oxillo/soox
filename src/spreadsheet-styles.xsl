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
  
  
</xsl:stylesheet>