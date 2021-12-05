<xsl:stylesheet
  version="3.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
  xmlns:soox="simple-office-open-xml"
  xmlns:sooxns="simple-office-open-xml/namespaces"
  exclude-result-prefixes="#all">
  
  <xd:doc scope="stylesheet">
    <xd:desc>
      <xd:p><xd:b>Created on:</xd:b> Nov 13, 2021</xd:p>
      <xd:p><xd:b>Author:</xd:b> Olivier XILLO</xd:p>
      <xd:p/>
    </xd:desc>
  </xd:doc>
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Generates the 'styles.xml' file information.</xd:p>
    </xd:desc>
    <xd:param name="simple_workbook"></xd:param>
  </xd:doc>
  <xsl:function name="soox:styles.xml" visibility="public">
    <xsl:param name="simple_workbook"/>
    <xsl:param name="filename"/>
    
    <xsl:variable name="content">
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <numFmts xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1">
          <numFmt numFmtId="164" formatCode="General"/>
        </numFmts>
        <fonts xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4">
          <font>
            <sz val="11"/>
            <color rgb="FF000000"/>
            <name val="Calibri"/>
            <family val="2"/>
            <charset val="238"/>
          </font>
          <font>
            <sz val="10"/>
            <name val="Arial"/>
            <family val="0"/>
          </font>
          <font>
            <sz val="10"/>
            <name val="Arial"/>
            <family val="0"/>
          </font>
          <font>
            <sz val="10"/>
            <name val="Arial"/>
            <family val="0"/>
          </font>
        </fonts>
        <fills xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">
          <fill>
            <patternFill patternType="none"/>
          </fill>
          <fill>
            <patternFill patternType="gray125"/>
          </fill>
        </fills>
        <borders xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1">
          <border diagonalUp="false" diagonalDown="false">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
          </border>
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
        <cellXfs xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">
          <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="false"
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
          </xf>
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
        <fills count="2">
          <fill>
            <patternFill patternType="none"/>
          </fill>
          <fill>
            <patternFill patternType="gray125"/>
          </fill>
        </fills>
        <borders count="1">
          <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
          </border>
        </borders>
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
    <xsl:sequence select="map{ $filename : map{
      'content': $content,
      'content-type':'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
      'relationship-type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
      }}"/>
  </xsl:function>
</xsl:stylesheet>