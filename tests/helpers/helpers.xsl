<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
    xmlns:xh="xspec-helpers"
    exclude-result-prefixes="xs xd"
    version="3.0">
    <xd:doc scope="stylesheet">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> Nov 13, 2021</xd:p>
            <xd:p><xd:b>Author:</xd:b> Olivier</xd:p>
            <xd:p></xd:p>
        </xd:desc>
    </xd:doc>
    
    <xsl:function name="xh:remove-whitespacesonly-nodes" as="document-node()">
        <xsl:param name="doc" as="document-node()"/>
        <xsl:apply-templates select="$doc" mode="xh:whitespacesonly-removal"/>
    </xsl:function>
    
    <xsl:template match="attribute() | document-node() | node()"
        mode="xh:whitespacesonly-removal">
        <xsl:copy>
            <xsl:apply-templates mode="#current" select="attribute() | node()" />
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="text()[normalize-space() => not()]"
        mode="xh:whitespacesonly-removal"/>
</xsl:stylesheet>