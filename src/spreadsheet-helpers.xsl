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
    <xsl:param name="num" as="xs:positiveInteger"/>
    
    <xsl:sequence select="$column-names[$num]"/>
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
    <xsl:param name="col" as="xs:positiveInteger"/>
    <xsl:param name="row" as="xs:integer"/>
    
    <xsl:sequence select="concat($column-names[$col],$row)"/>
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
  
  
  <xsl:variable name="letters" as="xs:string+" static="yes"
      select="('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z')"/>
  <xsl:variable name="two_letters" as="xs:string+" static="yes"
    select="$letters=>for-each(function($k) { $letters!concat($k,.) } )"/>
  <xsl:variable name="three_letters"  as="xs:string+" static="yes"
    select="$letters=>for-each(function($k) { $two_letters!concat($k,.) } )"/>
  <xsl:variable name="column-names" as="xs:string+" static="yes"
      select="($letters,$two_letters,$three_letters)=>subsequence(1,16384)"/>
  
  
  
    
</xsl:stylesheet>