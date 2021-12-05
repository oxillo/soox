
@set SOOX_HOME=%~dp0\..

@set CLASSPATH=%SOOX_HOME%\SaxonHE10-6J\saxon-he-10.6.jar;
@set CLASSPATH=%CLASSPATH%;%SOOX_HOME%\SaxonHE10-6J\soox-extensions.jar

::java net.sf.saxon.Transform  -it -config:../src/saxon-HE.config -xsl:test-extension.xsl -o:result.xml
java net.sf.saxon.Transform  -it -config:../src/saxon-HE.config -xsl:soox2OfficeOpenXml.xsl -o:result.xml