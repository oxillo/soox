
@set SOOX_HOME=%~dp0\..

@set CLASSPATH=%SOOX_HOME%\SaxonHE12-4J\saxon-he-12.4.jar;
@set CLASSPATH=%CLASSPATH%;%SOOX_HOME%\SaxonHE12-4J\soox-extensions.jar

::java net.sf.saxon.Transform  -it -config:../src/saxon-HE.config -xsl:test-extension.xsl -o:result.xml
java net.sf.saxon.Transform  -it -config:../src/saxon-HE.config -xsl:soox2OfficeOpenXml.xsl -o:result.xml