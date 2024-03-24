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
      <xd:p>For workbook and worksheets having a <style/> element,
        compute the new style (cascade) and apply it on dowstream elements.</xd:p>
      <xd:p>Rewrite element to remove attributes from <style/> element.</xd:p>
    </xd:desc>
    <xd:param name="inherited-style">The default style for workbook or the cascaded workbook style for worksheets</xd:param>
  </xd:doc>
  <xsl:template match="s:workbook[s:style]|s:worksheet[s:style]" mode="soox:spreadsheet-styles-cascade">
    <xsl:param name="inherited-style" tunnel="yes"/>
    
    <xsl:copy>
      <xsl:copy-of select="@*"/>
      <s:style>
        <xsl:apply-templates select="s:style/node()" mode="#current"/>
      </s:style>
      <xsl:apply-templates mode="#current" select="*[not(self::s:style)]">
        <xsl:with-param name="inherited-style" select="soox:cascade-style($inherited-style,s:style)" tunnel="yes"/>
      </xsl:apply-templates>
    </xsl:copy>
  </xsl:template>
  
  <xd:doc>
    <xd:desc>
      <xd:p>Cascade <style/> to <b>all</b> cell elements.</xd:p>
      <xd:p>(Re)write <style/> element to hold <b>all</b> style attributes with short notation expanded.</xd:p>
      <xd:p>compute and add style signatures to the <style/> element.</xd:p>
    </xd:desc>
    <xd:param name="inherited-style">The inherited style from worksheet</xd:param>
  </xd:doc>
  <xsl:template match="s:cell" mode="soox:spreadsheet-styles-cascade">
    <xsl:param name="inherited-style" tunnel="yes"/>
    
    <xsl:variable name="cascaded-style" select="if (s:style) then soox:cascade-style($inherited-style,s:style) else $inherited-style"/>
    <xsl:copy>
      <xsl:copy-of select="@*"/>
      <!-- Write the <style/> element with all attributes; computes signatures -->
      <xsl:element name="style" _namespace="soox">
        <!-- border style & color -->
        <!--xsl:variable name="border-signature" select="(
          $cascaded-style('border-left-style'),$cascaded-style('border-left-color'),
          $cascaded-style('border-right-style'),$cascaded-style('border-right-color'),
          $cascaded-style('border-top-style'),$cascaded-style('border-top-color'),
          $cascaded-style('border-bottom-style'),$cascaded-style('border-bottom-color'))=>string-join('#')"/-->
        <xsl:attribute name="border-signature" select="$cascaded-style('border-signature')"/>
        <!-- font -->
        <xsl:attribute name="font-signature" select="$cascaded-style('font-signature')"/> 
        <!-- fill -->
        <xsl:attribute name="fill-signature" select="$cascaded-style('fill-signature')"/>
        <!-- numeric format -->
        <xsl:attribute name="numeric-format-signature" select="$cascaded-style('numeric-format-signature')"/>
        <!-- style signature -->
        <!--xsl:attribute name="style-signature" select="($cascaded-style('border-signature'),$cascaded-style('font-signature'),$cascaded-style('fill-signature'),$cascaded-style('numeric-format-signature'))=>string-join(';')"/-->
        <xsl:attribute name="style-signature" select="$cascaded-style('style-signature')"/>
        <xsl:apply-templates select="s:style/node()" mode="#current"/>
      </xsl:element>
      <xsl:apply-templates mode="#current" select="* except s:style">
        <xsl:with-param name="inherited-style" select="$cascaded-style" tunnel="yes"/>
      </xsl:apply-templates>
    </xsl:copy>
  </xsl:template>
  
  
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Ignore the <style/> element as it is explicitly rewritten in other templates</xd:p>
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
    <xsl:param name="local" as="element(s:style)"/>
      
    <xsl:variable name="cascaded-style" select="$inherited=>
      soox:cascade-border-style($local)=>
      soox:cascade-fill-style($local)=>
      soox:cascade-font-style($local)=>
      soox:cascade-numeric-format($local)"/>
    <!--xsl:variable name="font-signature" select="($cascaded-style('font-family'),$cascaded-style('font-size'),$cascaded-style('font-weight'),$cascaded-style('font-style'))=>string-join('#')"/>
    <xsl:variable name="fill-signature" select="($cascaded-style('fill-style'),$cascaded-style('fill-color'))=>string-join('#')"/>
    <xsl:attribute name="numeric-format-signature" select="$cascaded-style('numeric-format')"/>
    <xsl:attribute name="style-signature" select="($cascaded-style('border-signature'),$font-signature,$fill-signature,$cascaded-style('numeric-format'))=>string-join(';')"/-->
    <xsl:sequence select="$cascaded-style=>map:put('style-signature',
      ($cascaded-style('border-signature'),$cascaded-style('font-signature'),$cascaded-style('fill-signature'),$cascaded-style('numeric-format-signature'))=>string-join(';'))"/>
  </xsl:function>
  
  
  
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
      <xsl:variable name="bordersTablemap" as="xs:string*"
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
    
    <xsl:variable name="signatures" as="xs:string*"
      select="distinct-values($styles/@style-signature)=>sort()"/>
    <xsl:map>
      <xsl:for-each select="$signatures">
        <xsl:map-entry key="current()" select="position()"/>
      </xsl:for-each>
    </xsl:map>
      
    
    <!--xsl:map>
      <xsl:for-each-group select="$styles" group-by="xs:string(@style-signature)">
        <xsl:sort order="ascending"/>
        <xsl:map-entry key="current-grouping-key()" select="position()"/>
      </xsl:for-each-group>
    </xsl:map-->
  </xsl:function>
  
 
  
  <xd:doc>
    <xd:desc>
      <xd:p>Compute the table of unique cell styles</xd:p>
    </xd:desc>
    <xd:param name="styles">The sequence of styles defined in the workbook</xd:param>
    <xd:param name="fontsTablemap"></xd:param>
    <xd:param name="fillsTablemap"></xd:param>
    <xd:param name="bordersTablemap"></xd:param>
    <xd:param name="numericFormatTableMap"></xd:param>
    <xd:return>A sequence of xf elements</xd:return>
  </xd:doc>
  <xsl:function name="soox:computeStyleCellXfsTable" as="element(sml:xf)*">
    <xsl:param name="styles" as="element(s:style)*"/>
    <xsl:param name="fontsTablemap" as="map(xs:string,element(sml:font))"/>
    <xsl:param name="fillsTablemap" as="map(xs:string,element(sml:fill))"/>
    <xsl:param name="borderSignatures" as="xs:string*"/>
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
    <xsl:variable name="fontSignatures" select="map:keys($fontsTablemap)=>sort()"/>
    <xsl:variable name="fillSignatures" select="map:keys($fillsTablemap)=>sort()"/>
    
    <xsl:for-each select="distinct-values($styles/@style-signature)">
      <xsl:variable name="tokens" select="current()=>tokenize(';')"/>
      <sml:xf>
        <xsl:attribute name="borderId" select="$borderSignatures=>index-of($tokens[1]) - 1"/>
        <xsl:attribute name="fontId" select="$fontSignatures=>index-of($tokens[2]) - 1"/>
        <xsl:attribute name="fillId" select="$fillSignatures=>index-of($tokens[3]) - 1"/>
        <xsl:attribute name="numFmtId" select="$numericFormatTableMap($tokens[4])"/>
        <!--xsl:attribute name="xfId" select="(0)[1]"/-->
        <xsl:attribute name="applyFont" select="'true'"/>
        <xsl:attribute name="applyBorder" select="'true'"/>
        <xsl:attribute name="applyAlignment" select="'false'"/>
        <xsl:attribute name="applyProtection" select="'false'"/>
        <!--alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false"
          indent="0" shrinkToFit="false"/>
        <protection locked="true" hidden="false"/-->
      </sml:xf>
    </xsl:for-each>
  </xsl:function>
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Define the default cell style. This is the top of the inheritance cascade</xd:p>
    </xd:desc>
  </xd:doc>
  <xsl:variable name="default-cell-style"
    select="map:merge((
    map {
    'border-left-style'  : 'none',
    'border-right-style' : 'none',
    'border-top-style'   : 'none',
    'border-bottom-style': 'none'
    },map {
    'border-left-color'  : soox:parse-color('black'),
    'border-right-color' : soox:parse-color('black'),
    'border-top-color'   : soox:parse-color('black'),
    'border-bottom-color': soox:parse-color('black'),
    'border-signature':'none#FF000000#none#FF000000#none#FF000000#none#FF000000'
    },map {
    'font-family' : 'Arial',
    'font-size'   : 12,
    'font-weight' : 'normal',
    'font-style'  : 'normal',
    'font-signature': 'Arial#12#normal#normal'
    },map {
    'fill-style'     : 'none',
    'fill-color'     : soox:parse-color('black'),
    'fill-signature' : 'none#FF000000' 
    },map {
    'numeric-format' : 'General',
    'numeric-format-signature' : 'General'
    },map {
    'style-signature' : 'none#FF000000#none#FF000000#none#FF000000#none#FF000000;Arial#12#normal#normal;none#FF000000;General'
    }))"/>
  
</xsl:stylesheet>