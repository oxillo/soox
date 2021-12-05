<xsl:stylesheet
  version="3.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
  xmlns:sml="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:soox="simplify-office-open-xml"
  xmlns:s="soox"
  exclude-result-prefixes="#all" extension-element-prefixes="soox">
  
  <xd:doc scope="stylesheet">
    <xd:desc>
      <xd:p><xd:b>Created on:</xd:b> Nov 21, 2021</xd:p>
      <xd:p><xd:b>Author:</xd:b> Olivier XILLO</xd:p>
      <xd:p>templates for conversion to/from Open Office</xd:p>
    </xd:desc>
  </xd:doc>
  
  
  <xsl:use-package name="soox:utils" version="1.0"/>
  
  <xsl:include href="spreadsheet-helpers.xsl"/>
  
  
  
  
  
  
  <!--=======================================================================================================-->
  <!--== Conversion templates to Open Office                                                               ==-->
  <!--=======================================================================================================-->
  <!--==                                                                                                   ==-->
  
  <xsl:mode name="soox:toOfficeOpenXml" visibility="public"/>
  <xsl:variable name="ns-sml" select="'http://schemas.openxmlformats.org/spreadsheetml/2006/main'" static="true"/>
  
  <xd:doc>
    <xd:desc>
      <xd:p></xd:p>
    </xd:desc>
  </xd:doc>
  <xsl:template match="s:worksheet" mode="soox:toOfficeOpenXml">
    <xsl:element name="worksheet" _namespace="{$ns-sml}">
      
      <!--sheetPr xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" filterMode="false">
        <pageSetUpPr fitToPage="false"/>
      </sheetPr-->
      
      <!-- compute the cell range min row/col-> max row/col -->
      <xsl:if test="descendant::s:cell"> 
        <xsl:element name="dimension" _namespace="{$ns-sml}">
          <xsl:variable name="all-rows" select="distinct-values(descendant::s:cell/@row)"/>
          <xsl:variable name="first-row" select="min($all-rows)"/>
          <xsl:variable name="last-row" select="max($all-rows)"/>
          <xsl:variable name="all-columns" select="distinct-values(descendant::s:cell/@col)"/>
          <xsl:variable name="first-column" select="min($all-columns) => xs:integer() => soox:convertColumnNumberToLetters()"/>
          <xsl:variable name="last-column" select="max($all-columns) => xs:integer() => soox:convertColumnNumberToLetters()"/>
          <xsl:attribute name="ref" select="concat($first-column, $first-row, ':', $last-column,$last-row)"/>
        </xsl:element>
      </xsl:if>
      
      <!--xsl:element name="sheetViews" _namespace="{$ns-sml}">
        <xsl:element name="sheetView" _namespace="{$ns-sml}">
          <xsl:attribute name="tabSelected">0</xsl:attribute>
          <xsl:attribute name="workbookViewId">0</xsl:attribute>
        </xsl:element>
      </xsl:element>
      
      <xsl:element name="sheetFormatPr" _namespace="{$ns-sml}">
        <xsl:attribute name="baseColWidth">10</xsl:attribute>
        <xsl:attribute name="defaultRowHeight">14.25</xsl:attribute>
      </xsl:element-->
      
      <xsl:apply-templates mode="#current"/>
      
      <!--xsl:element name="pageMargins" namespace="{$ns-sml}">
        <xsl:attribute name="left">0.7</xsl:attribute>
        <xsl:attribute name="right">0.7</xsl:attribute>
        <xsl:attribute name="top">0.75</xsl:attribute>
        <xsl:attribute name="bottom">0.75</xsl:attribute>
        <xsl:attribute name="header">0.3</xsl:attribute>
        <xsl:attribute name="footer">0.3</xsl:attribute>
      </xsl:element-->
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="s:data"  mode="soox:toOfficeOpenXml">
    
    <!--cols xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <col min="1" max="1" width="20.69921875" customWidth="1"/>
      <col min="2" max="2" width="70.69921875" customWidth="1"/>
    </cols-->
    <xsl:element name="sheetData" _namespace="{$ns-sml}">
      <xsl:for-each-group select="s:cell" group-by="@row" >
        <xsl:sort select="current-grouping-key()" order="ascending"/>
        <xsl:element name="row" _namespace="{$ns-sml}">
          <xsl:attribute name="r" select="current-grouping-key()"/>
          <!--xsl:attribute name="spans">1:2</xsl:attribute-->
          <xsl:attribute name="customFormat">false</xsl:attribute>
          <xsl:attribute name="ht">12.8</xsl:attribute>
          <xsl:attribute name="hidden">false</xsl:attribute>
          <xsl:attribute name="customHeight">false</xsl:attribute>
          <xsl:attribute name="outlineLevel">0</xsl:attribute>
          <xsl:attribute name="collapsed">false</xsl:attribute>
          <xsl:apply-templates select="current-group()" mode="#current"/>
        </xsl:element>
      </xsl:for-each-group>
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="s:cell"  mode="soox:toOfficeOpenXml">
    <xsl:param name="shared-strings" as="xs:string*" tunnel="yes"/>
    
    <xsl:variable name="option-date-as-number" select="true()"/>
    
    <xsl:variable name="cell-value" select="./s:v/text()"/>
    <xsl:variable name="has-value" select="string-length($cell-value) gt 0"/>
    <xsl:variable name="cell-formula" select="./s:f/text()"/>
    <xsl:variable name="has-formula" select="string-length($cell-formula) gt 0"/>
    <xsl:variable name="is-number" select="number($cell-value) ne xs:double('NaN')"/>
    <xsl:variable name="is-date" select="$cell-value castable as xs:date"/>
    
    <!-- Check if the value is defined in shared-strings and, if so , get its index -->
    <xsl:variable name="shared-string-index" select="if ($has-value) then index-of($shared-strings, $cell-value) else 0"/>
    
    <xsl:element name="c" _namespace="{$ns-sml}">
      <xsl:attribute name="r" select="soox:encode-cell-address(@col,@row)"/>
      <xsl:attribute name="s">
        <xsl:choose>
          <xsl:when test="$is-date and $option-date-as-number">1</xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <!-- content type of cell: 'b' for boolean, 'd' for date in ISO 8601, 'e' for error, 
           'inlineStr' for rich Text string, 's' for shared string, 'n' for number, 'str' for formula string 
           cf § 18.18.11 ST_CellType (Cell Type) -->
      <xsl:attribute name="t">
        <xsl:choose>
          <xsl:when test="$is-date and $option-date-as-number">n</xsl:when>
          <xsl:when test="$is-date">d</xsl:when>
          <xsl:when test="$shared-string-index">s</xsl:when>
          <xsl:when test="$is-number">n</xsl:when>
          <xsl:otherwise>inlineString</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      
      <xsl:if test="$has-formula">
        <xsl:element name="f" _namespace="{$ns-sml}">
          <xsl:value-of select="$cell-formula"/>
        </xsl:element>
      </xsl:if>
      <xsl:element name="v" _namespace="{$ns-sml}">
        <xsl:choose>
          <xsl:when test="count($shared-string-index) ge 1">
            <xsl:value-of select="$shared-string-index[1] - 1"/>
          </xsl:when>
          <xsl:when test="$is-date and $option-date-as-number">
            <xsl:value-of select="days-from-duration(xs:date($cell-value) - xs:date('1900-01-01')) + 2"/>
            <!-- +2 is for libreoffice not having the same reference as Excel -->
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$cell-value"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:element>
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="s:v"  mode="soox:toOfficeOpenXml" xml:space="default">
    <xsl:param name="shared-strings" as="xs:string*" tunnel="yes"/>
    
    <xsl:variable name="option-date-as-number" select="true()"/>
    
    
    <xsl:variable name="is-number" select="number(text()) ne xs:double('NaN')"/>
    <xsl:variable name="is-date" select="text() castable as xs:date"/>
    
    <!-- Check if the value is defined in shared-strings and, if so , get its index -->
    <xsl:variable name="shared-string-index" select="index-of($shared-strings,text())"/>
    
    <xsl:element name="c" _namespace="{$ns-sml}">
      <xsl:attribute name="r" select="soox:convertColumnNumberToLetters(ancestor::s:cell/@col)||ancestor::s:cell/@row"/>
      <xsl:attribute name="s">
        <xsl:choose>
          <xsl:when test="$is-date and $option-date-as-number">1</xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <!-- content type of cell: 'b' for boolean, 'd' for date in ISO 8601, 'e' for error, 
           'inlineStr' for rich Text string, 's' for shared string, 'n' for number, 'str' for formula string 
           cf § 18.18.11 ST_CellType (Cell Type) -->
      <xsl:attribute name="t">
        <xsl:choose>
          <xsl:when test="$is-date and $option-date-as-number">n</xsl:when>
          <xsl:when test="$is-date">d</xsl:when>
          <xsl:when test="$shared-string-index">s</xsl:when>
          <xsl:when test="$is-number">n</xsl:when>
          <xsl:otherwise>inlineString</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      
      <xsl:element name="v" _namespace="{$ns-sml}">
        <xsl:variable name="shared-string-index" select="index-of($shared-strings,text())"/>
        <xsl:choose>
          <xsl:when test="count($shared-string-index) ge 1">
            <xsl:value-of select="$shared-string-index[1] - 1"/>
          </xsl:when>
          <xsl:when test="$is-date and $option-date-as-number">
            <xsl:value-of select="days-from-duration(xs:date(text()) - xs:date('1900-01-01')) + 2"/>
            <!-- +2 is for libreoffice not having the same reference as Excel -->
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="text()"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:element>  
    </xsl:element>  
  </xsl:template>
  
  <xsl:template match="s:f"  mode="soox:toOfficeOpenXml" xml:space="default">
    <xsl:element name="f" _namespace="{$ns-sml}">
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="text()" mode="soox:toOfficeOpenXml"/>
  
  
  
  <!--=======================================================================================================-->
  <!--== Conversion templates to Open Office                                                               ==-->
  <!--=======================================================================================================-->
  <!--==                                                                                                   ==-->
  <xsl:mode name="soox:fromOfficeOpenXml" visibility="public"/>
  <xsl:variable name="ns-soox" select="'soox'" visibility="private"/>
  
  
  <xsl:template match="/" mode="soox:fromOfficeOpenXml">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>
  
  <xsl:template match="sml:workbook" mode="soox:fromOfficeOpenXml">
    <xsl:element name="workbook" namespace="soox">
      <xsl:apply-templates mode="#current"/>
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="sml:sheet" mode="soox:fromOfficeOpenXml">
    <xsl:param name="file-hierarchy" tunnel="yes"/>
    <xsl:param name="relationship" tunnel="yes"/>
    <xsl:param name="base" tunnel="yes"/>
    
    <xsl:variable name="relation-id" select="@*:id" as="xs:string"/>
    <xsl:variable name="worksheet">
      <xsl:variable name="fname" select="$relationship//*[@Id = $relation-id]/@Target"/>
      <xsl:sequence select="$file-hierarchy => soox:get-content( $base||$fname )"/>
    </xsl:variable>
    <xsl:apply-templates select="$worksheet" mode="soox:fromOfficeOpenXml">
      <xsl:with-param name="sheet-attributes" select="map{'name':@name, 'state':@state}" tunnel="yes"/>
    </xsl:apply-templates>
  </xsl:template>
  
  <xsl:template match="sml:worksheet" mode="soox:fromOfficeOpenXml">
    <xsl:param name="sheet-attributes" tunnel="yes"/>
    
    <xsl:element name="worksheet" namespace="soox">
      <xsl:attribute name="name" select="$sheet-attributes?name"/>
      <xsl:apply-templates mode="#current"/>
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="sml:sheetData" mode="soox:fromOfficeOpenXml">
    <xsl:element name="data" namespace="soox">
      <xsl:apply-templates mode="#current"/>
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="sml:*" mode="soox:fromOfficeOpenXml">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>
  
  <xsl:template match="sml:c" mode="soox:fromOfficeOpenXml">
    <xsl:element name="cell" namespace="soox">
      <xsl:variable name="address" select="soox:decode-cell-address(@r)"/>
      <xsl:attribute name="col" select="$address[1]"/>
      <xsl:attribute name="row" select="$address[2]"/>
      <xsl:apply-templates mode="#current"/>    
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="sml:v" mode="soox:fromOfficeOpenXml">
    <xsl:param name="sharedstrings-list" tunnel="yes"/>
    
    <xsl:element name="v" namespace="soox">
      <xsl:choose>
        <xsl:when test="../@t = 's'">
          <xsl:variable name="sharedstrings-index" select="text() + 1"/> 
          <xsl:value-of select="$sharedstrings-list[ $sharedstrings-index ]"/>
        </xsl:when>
        <xsl:otherwise><xsl:value-of select="text()"/></xsl:otherwise>
      </xsl:choose>
    </xsl:element>
  </xsl:template>
  
  <xd:doc>
    <xd:desc>
      <xd:p>Copy formula as is in soox</xd:p>
    </xd:desc>
  </xd:doc>
  <xsl:template match="sml:f" mode="soox:fromOfficeOpenXml">
    <xsl:element name="f" namespace="soox">
      <xsl:value-of select="text()"/>
    </xsl:element>
  </xsl:template>
    
</xsl:stylesheet>