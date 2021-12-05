<?xml version="1.0" encoding="UTF-8"?>
<xsl:package
    name="soox:utils"
    package-version="1.0"
    version="3.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
    xmlns:map="http://www.w3.org/2005/xpath-functions/map"
    xmlns:arch="http://expath.org/ns/archive"
    xmlns:file="http://expath.org/ns/file"
    xmlns:bin="http://expath.org/ns/binary"
    xmlns:soox="simplify-office-open-xml"
    xmlns:sooxns="simplify-office-open-xml/namespaces"
    exclude-result-prefixes="#all">
    
    
    <xsl:output name="standalone" method="xml" standalone="yes" indent="yes"/>
    
    <xsl:variable name="sooxns:sml" select="'http://schemas.openxmlformats.org/spreadsheetml/2006/main'" visibility="final"/>
    <xsl:variable name="sooxns:rels" select="'http://schemas.openxmlformats.org/package/2006/relationships'" visibility="final"/>
    
    
    
    <xsl:function name="soox:build-relationship-file" visibility="public">
        <xsl:param name="part-files"/>
        <xsl:param name="part-prefix"/>
        
        
        <xsl:variable name="content">  
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <xsl:for-each select="map:keys($part-files)">
                    <xsl:sort select="current()"/>
                    <Relationship>
                        <xsl:attribute name="Id" select="'rId'||position()"/>
                        <xsl:attribute name="Type" select="$part-files(current())('relationship-type')"/>
                        <xsl:attribute name="Target" select="substring-after(current(),$part-prefix || '/')"/>
                    </Relationship>
                </xsl:for-each>
            </Relationships>
        </xsl:variable>
        
        <xsl:sequence select="map{
            'content': $content,
            'content-type':'application/vnd.openxmlformats-package.relationships+xml'
            }"/>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Extract an XML file from a file hierarchy</xd:p>
        </xd:desc>
        <xd:param name="file-hierarchy"></xd:param>
        <xd:param name="xmlfile">the name of the XML file to extract</xd:param>
        <xd:return>an XML document-node</xd:return>
    </xd:doc>
    <xsl:function name="soox:extract-xmlfile-from-file-hierarchy" visibility="final">
        <xsl:param name="file-hierarchy" as="map(xs:string,map(xs:string,item()*))"/>
        <xsl:param name="xmlfile"/>
        
        <xsl:sequence select="$file-hierarchy($xmlfile)('content') => bin:decode-string('UTF-8') => parse-xml()"/>
    </xsl:function>
    
    <xsl:function name="soox:get-relationships-of-type" visibility="public">
        <xsl:param name="elem" as="document-node()"/>
        <xsl:param name="type" as="xs:string"/>
        
        <xsl:sequence select="$elem//Q{http://schemas.openxmlformats.org/package/2006/relationships}Relationship[@Type=$type]"/>
    </xsl:function>

    
</xsl:package>