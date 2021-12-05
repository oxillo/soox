<xsl:stylesheet
  version="3.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
  xmlns:sml="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:soox="simple-office-open-xml"
  xmlns:s="soox"
  exclude-result-prefixes="#all" extension-element-prefixes="soox">
  
  <xd:doc scope="stylesheet">
    <xd:desc>
      <xd:p><xd:b>Created on:</xd:b> Nov 11, 2021</xd:p>
      <xd:p><xd:b>Author:</xd:b> Olivier XILLO</xd:p>
      <xd:p>Helper functions for spreadsheets</xd:p>
    </xd:desc>
  </xd:doc>
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Convert the column number <xd:i>num</xd:i>(1-based) to columns letter 'A'</xd:p>
    </xd:desc>
    <xd:param name="num">column number (integer greater or equal than 1)</xd:param>
    <xd:return>a string for A1 cell address</xd:return>
  </xd:doc>
  <xsl:function name="soox:convertColumnNumberToLetters" as="xs:string" visibility="private">
    <xsl:param name="num" as="xs:integer"/>
    
    <!-- if we just do $num mod 26, we will get 1 for A, 25 for Y and 0 for Z
      so we do $num-1 mod 26 to have 0 for A and 25 for Z -->
    <xsl:variable name="letter" select="codepoints-to-string((65 + (($num - 1) mod 26)))"/>
    <xsl:choose>
      <!-- The number is between 1 and 26, return the associated letter 'A'..'Z' -->
      <xsl:when test="$num le 0">
        <!--xsl:sequence select="error(QName('http://www.w3.org/2005/xqt-errors'),'error description of my template',
        ('error', 'object', 'of', 'my', 'template'))" /-->
        <xsl:message terminate="yes">soox:convertColumnNumberToLetters was called with an invalid argument(<xsl:value-of select="$num"/>)</xsl:message>
      </xsl:when>
      <!-- The number is between 1 and 26, return the associated letter 'A'..'Z' -->
      <xsl:when test="$num le 26">
        <xsl:sequence select="$letter"/>
      </xsl:when>
      <!-- For number above 26, recursively call the function for preceding letters --> 
      <xsl:otherwise>
        <xsl:sequence select="concat(soox:convertColumnNumberToLetters($num idiv 26),$letter)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
  <xd:doc>
    <xd:desc>
      <xd:p>Encode a cell address into A1 format</xd:p>
    </xd:desc>
    <xd:param name="col">an integer for the column (first column is 1)</xd:param>
    <xd:param name="row">an integer for the row (first row is 1)</xd:param>
    <xd:return>a string in A1 format or the empty string ''</xd:return>
  </xd:doc>
  <xsl:function name="soox:encode-cell-address" visibility="public">
    <xsl:param name="col" as="xs:integer"/>
    <xsl:param name="row" as="xs:integer"/>
    
    <xsl:sequence select="soox:convertColumnNumberToLetters($col)||$row"/>
  </xsl:function>
  
  <xd:doc>
    <xd:desc>
      <xd:p>Encode a cell address into A1 format</xd:p>
    </xd:desc>
    <xd:param name="col-row">a sequence of 2 (or more) integers. The first is the row, the second the column. The remaining ones are ignored.</xd:param>
    <xd:return>a string in A1 format or the empty string ''</xd:return>
  </xd:doc>
  <xsl:function name="soox:encode-cell-address" visibility="public">
    <xsl:param name="col-row" as="xs:integer+"/>
    
    <xsl:sequence select="if (count($col-row) ge 2) then soox:encode-cell-address($col-row[1],$col-row[2]) else ''"/>
  </xsl:function>
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Decode a cell address in A1 format to a couple of integer</xd:p>
    </xd:desc>
    <xd:param name="address">a string in A1 format</xd:param>
    <xd:return>a sequence of 2 integers or the empty sequence</xd:return>
  </xd:doc>
  <xsl:function name="soox:decode-cell-address" visibility="public">
    <xsl:param name="address" as="xs:string"/>
    
    <xsl:analyze-string select="$address => upper-case()" regex="([A-Z]+)([1-9][0-9]*)">
      <xsl:matching-substring>
        <xsl:variable name="column" select="regex-group(1) => string-to-codepoints() => fold-left(0, 
          function($result,$e){ $result * 26 + ($e - 64) })" as="xs:integer"/>
        <xsl:sequence select="($column,xs:integer(regex-group(2)))"/>
      </xsl:matching-substring>
    </xsl:analyze-string>
  </xsl:function>
  
  
  
    
</xsl:stylesheet>