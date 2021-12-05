<?xml version="1.0" encoding="UTF-8"?>
<xsl:package
    name="soox:opc"
    package-version="1.0"
    version="3.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
    xmlns:map="http://www.w3.org/2005/xpath-functions/map"
    xmlns:soox="simple-office-open-xml"
    xmlns:opc="simple-office-open-xml/open-packaging-conventions"
    exclude-result-prefixes="#all">
    <xd:doc scope="package">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> Dec 4, 2021</xd:p>
            <xd:p><xd:b>Author:</xd:b>Olivier XILLO</xd:p>
            <xd:p>SOOX package for Open Packaging Conventions
            </xd:p>
        </xd:desc>
    </xd:doc>
    
    <xsl:use-package name="soox:utils" version="1.0"/>
    
    <xd:doc>
        <xd:desc>
            <xd:p>Write the package (parts hierarchy) into a zip file</xd:p>
        </xd:desc>
        <xd:param name="parts">the set of parts as a map</xd:param>
        <xd:param name="uri">uri of the zip file to write</xd:param>
    </xd:doc>
    <xsl:function name="opc:zip" visibility="public">
        <xsl:param name="parts" as="map(*)"/>
        <xsl:param name="uri"/>
        
        <xsl:for-each select="map:keys($parts)">
            <xsl:message select="current()"></xsl:message>
        </xsl:for-each>
        
        
        <!-- Generate [Content-Types].xml from parts -->
        <xsl:variable name="package" select="map:merge(($parts,opc:generate-content-types($parts),opc:generate-package-relationships($parts)))"/>
        
        <!-- function to encode each XML file/element into a UTF-8 string -->
        <xsl:variable name="encode-file-entry-fn" select="function($fname, $fmap){
            map:entry(substring-after($fname,'/'), map{'content' : $fmap('content') => serialize() => bin:encode-string('UTF-8')})
            }" xmlns:bin="http://expath.org/ns/binary"/>
        
        <!-- encode each file of the hierarchy and build the archive -->
        <xsl:variable name="zipped-file-hierarchy"
            select="map:for-each( $package, $encode-file-entry-fn ) => map:merge() => arch:create-map()" xmlns:arch="http://expath.org/ns/archive"/>
        
        <!-- write the archive to disk -->
        <xsl:sequence select="file:write-binary( $uri, $zipped-file-hierarchy )" xmlns:file="http://expath.org/ns/file"/>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Read a package from a zip file</xd:p>
        </xd:desc>
        <xd:param name="uri">uri of the file to read from</xd:param>
    </xd:doc>
    <xsl:function name="opc:unzip" visibility="public">
        <xsl:param name="uri"/>
        
        <!-- read the archive from disk -->
        <xsl:variable  name="unzipped" select="file:read-binary( $uri )" xmlns:file="http://expath.org/ns/file"/>
        
        <!-- -->
        <xsl:variable  name="parts" select="arch:entries-map( $unzipped, true() )"  xmlns:arch="http://expath.org/ns/archive"/>
        
        <xsl:variable name="content-types" select="$parts => soox:extract-xmlfile-from-file-hierarchy('[Content_Types].xml')"/>
        <xsl:variable name="overrides" select="$content-types//Override" xpath-default-namespace="http://schemas.openxmlformats.org/package/2006/content-types"/>
        <xsl:variable name="defaults" select="$content-types//Default" xpath-default-namespace="http://schemas.openxmlformats.org/package/2006/content-types"/>
        
        <xsl:map>
            <xsl:for-each select="map:keys($parts)[. ne '[Content_Types.xml']">
                <xsl:variable name="partname" select="current()"/>
                <xsl:variable name="type" select="($overrides[@PartName=$partname]/@ContentType,$defaults[ends-with($partname,@Extension)]/@ContentType,'')[1]" xpath-default-namespace="http://schemas.openxmlformats.org/package/2006/content-types"/>
                <xsl:map-entry key="$partname">
                    <xsl:map>
                        <xsl:map-entry key="'content'" select="$parts($partname)('content')"/>
                        <xsl:map-entry key="'content-type'" select="$type"/>
                    </xsl:map>
                </xsl:map-entry>
            </xsl:for-each>    
        </xsl:map>
        
        
        <!--xsl:sequence select="if(map:contains($parts,'[Content_Types].xml')) then $parts else map{}"/-->
    </xsl:function>
    
    
    
    <xd:doc>
        <xd:desc>Generates the content of [Content_Types].xml from parts</xd:desc>
        <xd:param name="parts">The list of files to include in [Content_Types].xml</xd:param>
    </xd:doc>
    <xsl:function name="opc:generate-content-types" visibility="private">
        <xsl:param name="parts" as="map(*)"/>
        
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
                <xsl:for-each select="map:keys($parts)">
                    <xsl:variable name="file-name" select="current()"/>
                    <xsl:variable name="file-content-type" select="$parts(current())('content-type')"/>
                    <xsl:for-each select="map:keys($default-extensions)">
                        <xsl:variable name="extension" select="current()"/>
                        <xsl:if test="$file-name => ends-with('.'||$extension)">
                            <xsl:if test="$file-content-type ne $default-extensions($extension)">
                                <Override xmlns="http://schemas.openxmlformats.org/package/2006/content-types"
                                    PartName="{$file-name}" ContentType="{$file-content-type}"/>
                            </xsl:if>
                        </xsl:if>
                    </xsl:for-each>
                </xsl:for-each>
            </Types>
        </xsl:variable>
        
        <!-- return a map -->
        <xsl:sequence select="map{'/[Content_Types].xml' : map{
            'content':$content}}"/>
    </xsl:function>
    
    
    <xd:doc>
        <xd:desc>Generates the package relationships from parts</xd:desc>
        <xd:param name="parts">The list of parts</xd:param>
    </xd:doc>
    <xsl:function name="opc:generate-package-relationships" visibility="private">
        <xsl:param name="parts" as="map(*)"/>
        
        <!-- retrieve @Target from every relationship parts (.rels) -->
        <xsl:variable name="relationship-targets" as="xs:string*">
            <xsl:for-each select="map:keys($parts)[ends-with(.,'.rels') and starts-with(.,'/')]">
                <xsl:variable name="base" select="substring-before(current(),'/_rels/')||'/'"/>
                <xsl:sequence select="distinct-values($parts(current())('content')//Relationship/@Target) ! concat($base,.)" xpath-default-namespace="http://schemas.openxmlformats.org/package/2006/relationships"/>
            </xsl:for-each>
        </xsl:variable>
        <xsl:message select="string-join($relationship-targets,'||')"></xsl:message>
        
        <xsl:map>
            <xsl:map-entry key="'/_rels/.rels'">
                <xsl:map>
                    <xsl:map-entry key="'content'">
                        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                            <xsl:for-each select="map:keys($parts)[not(ends-with(.,'.rels')) and starts-with(.,'/') and not(.=$relationship-targets)]">
                                <Relationship Id="rId{position()}" Type="{$parts(current())('relationship-type')}" Target="{substring-after(current(),'/')}"/>
                            </xsl:for-each>
                        </Relationships>
                    </xsl:map-entry>
                </xsl:map>
            </xsl:map-entry>
        </xsl:map>
    </xsl:function>
</xsl:package>
