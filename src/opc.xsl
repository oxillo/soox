<?xml version="1.0" encoding="UTF-8"?>
<xsl:package
    name="soox:opc"
    package-version="1.0"
    version="3.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
    xmlns:map="http://www.w3.org/2005/xpath-functions/map"
    xmlns:soox="simplify-office-open-xml"
    xmlns:opc="simplify-office-open-xml/open-packaging-conventions"
    xmlns:bin="http://expath.org/ns/binary"
    xmlns:arch="http://expath.org/ns/archive"
    xmlns:file="http://expath.org/ns/file"
    exclude-result-prefixes="#all"
    extension-element-prefixes="bin arch file soox">
    <xd:doc scope="package">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> Dec 4, 2021</xd:p>
            <xd:p><xd:b>Author:</xd:b>Olivier XILLO</xd:p>
            <xd:p>SOOX package for Open Packaging Conventions
            </xd:p>
        </xd:desc>
    </xd:doc>
    
    <xsl:use-package name="soox:utils" version="1.0"/>
    
    <xsl:variable name="need-soox-extensions" static="true" 
        select="not(function-available('bin:encode-string') and function-available('arch:create-map') and function-available('file:write-binary'))"/>
    
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
        
        
        <xsl:variable name="serialized-package" as="map(*)">
            <xsl:map>
                <xsl:for-each select="map:keys($package)">
                    <xsl:variable name="content" select="$package(current())('content')=>serialize()"/>
                    <xsl:map-entry key="substring-after(current(),'/')" select="map{'content':$content=> bin:encode-string('UTF-8')}" use-when="not($need-soox-extensions)"/>
                    <xsl:map-entry key="substring-after(current(),'/')" select="map{'content':$content}" use-when="$need-soox-extensions"/>
                </xsl:for-each>
            </xsl:map>
        </xsl:variable>

        <!-- write the archive to disk -->
        <xsl:sequence select="file:write-binary( $uri, arch:create-map($serialized-package))" use-when="not($need-soox-extensions)"/>
        <xsl:sequence select="soox:zip( $serialized-package, map{'uri':substring-after($uri,'file:/')})" use-when="$need-soox-extensions"/>
    </xsl:function>
    
    
    
    
    
    <xd:doc>
        <xd:desc>
            <xd:p>Read a package from a zip file</xd:p>
        </xd:desc>
        <xd:param name="uri">uri of the file to read from</xd:param>
    </xd:doc>
    <xsl:function name="opc:unzip" visibility="public">
        <xsl:param name="uri"/>
        
        <xsl:variable name="parts" as="map(*)" use-when="not($need-soox-extensions)">
            <!-- read the archive from disk -->
            <xsl:variable  name="unzipped" select="file:read-binary( $uri )"/>
            
            <!-- -->
            <xsl:variable  name="encoded-file-hierarchy" select="arch:entries-map( $unzipped, true() )"/>
            <xsl:variable name="decode" select="function($k,$v){
                if(ends-with($k,'.xml') or ends-with($k,'.rels')) then 
                map:entry('/'||$k, map{'content': parse-xml(bin:decode-string($v?content,'UTF-8'))})
                else
                map:entry('/'||$k, map{'content': $v?content})
                }" use-when="function-available('bin:decode-string')"/>
            <xsl:sequence select="map:for-each($encoded-file-hierarchy,$decode)=>map:merge()"/>  
        </xsl:variable>
        
        <xsl:variable name="parts" as="map(*)" use-when="$need-soox-extensions">
            <xsl:variable name="unzipped" select="soox:unzip(substring-after($uri,'file:/'))"/>
            <xsl:map>
                <xsl:for-each select="map:keys($unzipped)">
                    <xsl:variable name="fname" select="current()"/>
                    <xsl:variable name="content" select="$unzipped($fname)('content')"/>
                    <xsl:variable name="isxml" select="ends-with($fname,'.xml') or ends-with($fname,'.rels')"/>
                    <xsl:map-entry key="'/'||$fname" select="map{'content' : if ($isxml) then parse-xml($content) else $content}"/>
                </xsl:for-each>  
            </xsl:map>
        </xsl:variable>
        
        
        <xsl:variable name="content-types" select="$parts=>soox:get-content('/[Content_Types].xml')"/>
        <xsl:variable name="overrides" select="$content-types//Override" xpath-default-namespace="http://schemas.openxmlformats.org/package/2006/content-types"/>
        <xsl:variable name="defaults" select="$content-types//Default" xpath-default-namespace="http://schemas.openxmlformats.org/package/2006/content-types"/>
        
        <xsl:map>
            <xsl:for-each select="map:keys($parts)[. ne '/[Content_Types.xml']">
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
