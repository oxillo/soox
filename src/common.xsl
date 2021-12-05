<xsl:package
  name="soox:common"
  package-version="1.0"
  version="3.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
  xmlns:sooxns="simplify-open-office-xml/namespaces"
  xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  xmlns:soox="simplify-office-open-xml"
  xmlns:s="soox"
  xmlns:bin="http://expath.org/ns/binary"
  exclude-result-prefixes="#all"
  extension-element-prefixes="bin">
  
  
  <xsl:use-package name="soox:spreadsheet" version="1.0"/>
  <xsl:use-package name="soox:utils" version="1.0"/>
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Extract the file hierarchy that corresponds to the office part </xd:p>
    </xd:desc>
    <xd:param name="file-hierarchy">The Open Office file-hierarchy</xd:param>
  </xd:doc>
  <xsl:function name="soox:extract-office-part" as="map(xs:string, map(xs:string, item()*))" visibility="public">
      <xsl:param name="file-hierarchy" as="map(xs:string, map(xs:string, item()*))"/>

      
      <xsl:variable name="top-relationship" select="$file-hierarchy => soox:extract-xmlfile-from-file-hierarchy('_rels/.rels')"/>
      <xsl:variable name="office-fname" select="($top-relationship => soox:get-relationships-of-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'))[1]/@Target"/>
      <xsl:message expand-text="true">Trying to decode workbook {$office-fname}</xsl:message>
      
      <xsl:variable name="fname" select="tokenize($office-fname, '/')[last()]"/>
      <xsl:variable name="remapped-file-hierarchy" as="map(*)">
        <xsl:variable name="remap-fn" as="function(*)">
          <xsl:variable name="prefix" select="(tokenize($office-fname, '/')[position() lt last()]=>string-join('/')) || '/'"/>
          <xsl:sequence select="function($k,$file){
            if ($k=>starts-with($prefix)) then
              map:entry($k=>substring-after($prefix), $file)
            else ()}"/>
        </xsl:variable>
        <xsl:sequence select="$file-hierarchy => map:for-each($remap-fn) => map:merge()"/>
      </xsl:variable>
      <!-- Add a special key '$root'->'name'->workbook.xml-->
      <xsl:sequence select="$remapped-file-hierarchy => map:put('$root',map{'name':$fname})"/>
  </xsl:function>
  
</xsl:package>