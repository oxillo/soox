<xsl:package
  name="soox:common"
  package-version="1.0"
  version="3.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
  xmlns:sooxns="simple-open-office-xml/namespaces"
  xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  xmlns:soox="simple-open-office-xml"
  xmlns:opc="simple-open-office-xml/open-packaging-conventions"
  xmlns:s="soox"
  xmlns:bin="http://expath.org/ns/binary"
  exclude-result-prefixes="#all"
  extension-element-prefixes="bin opc">
  
  
  <xsl:use-package name="soox:opc" version="1.0"/>
  <xsl:use-package name="soox:spreadsheet" version="1.0"/>
  <xsl:use-package name="soox:utils" version="1.0"/>
  
  
  <xsl:variable name="soox:ns" select="map{
    'relationship':'http://schemas.openxmlformats.org/package/2006/relationships'}" visibility="final"/>
  
  
  
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Generate the Open Office file hierarchy from the input SOOX document</xd:p>
    </xd:desc>
    <xd:param name="doc">the SOOX document</xd:param>
    <xd:param name="options"></xd:param>
    <xd:return>A map ( filepath : file details (content, type,...) that represents the file hierachy of OpenOffice document</xd:return>
  </xd:doc>
  <xsl:function name="soox:generate-OpenOffice-file-hierarchy" visibility="public">
    <xsl:param name="doc" as="document-node()"/>
    <xsl:param name="options"/>
      
    <xsl:variable name="parts" as="map(*)*">
      <xsl:variable name="soox-root-element" select="$doc/element()"/>
      <xsl:choose>
        <xsl:when test="node-name($soox-root-element) eq QName('soox','workbook')">
          <xsl:sequence select="$soox-root-element=>soox:encode-spreadsheet_part()"/>
        </xsl:when>
        <xsl:otherwise/>
      </xsl:choose>
    </xsl:variable>
    
    <xsl:sequence select="map:merge($parts)"/>
  </xsl:function>
  
  
  
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Write an OpenOffice file hierarchy as a zip file</xd:p>
    </xd:desc>
    <xd:param name="file-hierarchy">the set of OpenOffice files as a map</xd:param>
    <xd:param name="uri">uri of the file to write</xd:param>
  </xd:doc>
  <xsl:function name="soox:zip-OpenOffice-file-hierarchy" visibility="public">
    <xsl:param name="file-hierarchy" as="map(*)"/>
    <xsl:param name="uri"/>
    
    <!-- write the archive to disk -->
    <xsl:sequence select="$file-hierarchy=>opc:zip($uri)"/>
  </xsl:function>
  
  
  
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Read an OpenOffice file hierarchy from zip file</xd:p>
    </xd:desc>
    <xd:param name="uri">uri of the file to write</xd:param>
  </xd:doc>
  <xsl:function name="soox:unzip-OpenOffice-file-hierarchy" visibility="public">
    <xsl:param name="uri"/>
    
    <!-- unpack parts -->
    <xsl:variable  name="parts" select="opc:unzip($uri)"/>
    
    <xsl:if test="$parts=>map:contains('[Content_Types].xml')">
      <!-- decode [Content_Types].xml -->
      <xsl:variable name="content-types" select="$parts => soox:extract-xmlfile-from-file-hierarchy('[Content_Types].xml')"/>
      <xsl:variable name="workbook" select="$content-types/Types/Override[@ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml']/@PartName" xpath-default-namespace="http://schemas.openxmlformats.org/package/2006/content-types"/>
      <xsl:variable name="document" select="$content-types/Types/Override[@ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml']/@PartName" xpath-default-namespace="http://schemas.openxmlformats.org/package/2006/content-types"/>
      <xsl:if test="$workbook">
        <xsl:sequence select="$parts => soox:extract-office-part() => soox:decode-spreadsheet-part()"/>
      </xsl:if>
      <xsl:if test="$document">
        <xsl:message expand-text="true">Document is {$document}</xsl:message>
      </xsl:if>
    </xsl:if>
  </xsl:function>
  
  
  
  
  <xd:doc>
    <xd:desc>Generates the content of [Content_Types].xml</xd:desc>
    <xd:param name="files">The list of files to include in [Content_Types].xml</xd:param>
  </xd:doc>
  <xsl:function name="soox:content_types" visibility="public">
    <xsl:param name="files" as="map(*)"/>
    
    <xsl:variable name="default-extensions" select="map {
      'bin':'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings',
      'rels':'application/vnd.openxmlformats-package.relationships+xml',
      'xml':'application/xml'
      }"/>
    
    <xsl:variable name="content">
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <!-- Associate MIME types to default extensions -->
        <xsl:for-each select="map:keys($default-extensions)">
          <Default xmlns="http://schemas.openxmlformats.org/package/2006/content-types"
            Extension="{current()}" ContentType="{$default-extensions(current())}"/>
        </xsl:for-each>
        
        <!-- Override MIME types -->
        <xsl:for-each select="map:keys($files)">
          <xsl:variable name="file-name" select="current()"/>
          <xsl:variable name="file-content-type" select="$files(current())('content-type')"/>
          <xsl:for-each select="map:keys($default-extensions)">
            <xsl:variable name="extension" select="current()"/>
            <xsl:if test="$file-name => ends-with($extension)">
              <xsl:if test="$file-content-type ne $default-extensions($extension)">
                <Override xmlns="http://schemas.openxmlformats.org/package/2006/content-types"
                  PartName="/{$file-name}" ContentType="{$file-content-type}"/>
              </xsl:if>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </Types>
    </xsl:variable>
 
    <xsl:sequence select="map{'[Content_Types].xml' : map{
      'content':$content}}"/>
  </xsl:function>
  
  
  
  
  
  <xd:doc>
    <xd:desc>
      <xd:p>Generate the top-level relationship file depending on the </xd:p>
    </xd:desc>
    <xd:param name="doc"></xd:param>
  </xd:doc>
  <xsl:function name="soox:top-level-relationship" visibility="public">
    <xsl:param name="doc" as="document-node()"/>
    
    <xsl:variable name="property-files" select="map{
      'docProps/core.xml' : map{'relationship-type' : 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'},
      'docProps/app.xml' : map{'relationship-type' : 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/extended-properties'},
      'docProps/custom.xml' : map{'relationship-type' : 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/cocustomre-properties'}}"/>
    <xsl:variable name="main-part-filename">
      <xsl:choose>
        <xsl:when test="$doc/element()=>node-name() eq QName('soox','workbook')">/xl/workbook.xml</xsl:when>
        <xsl:otherwise>unknown</xsl:otherwise>
      </xsl:choose>  
    </xsl:variable>
    <xsl:variable name="main-part-file" select="map{
      $main-part-filename: map{'relationship-type' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'}}"/>
      
    <xsl:sequence select="map{'_rels/.rels' : soox:build-relationship-file(map:merge(($main-part-file,$property-files)),'')}"/>
  </xsl:function>
  
  
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