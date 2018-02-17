<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet exclude-result-prefixes="#all" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:tei="http://www.tei-c.org/ns/1.0"
    xmlns:ms="urn:schemas-microsoft-com:office:spreadsheet"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
    xmlns:html="http://www.w3.org/TR/REC-html40"
    version="2.0">
    <xsl:output omit-xml-declaration="yes"/>
    
    <xsl:template match="/">
<xsl:result-document href="spreadsheet.csv">"Repository","Title","Call No.","Language(s)","Date","CompuDate","Place","Keywords","Front Cover","OPenn Record","Data","Page Turning","Penn in Hand Record","Summary","Collation View","Online","Month Photographed"<xsl:for-each select="collection(iri-to-uri('../viewshare/TEI/?select=*.xml;recurse=yes'))">
            <xsl:variable name="institution" select="//tei:institution"/>
            <xsl:variable name="repository" select="//tei:repository"/>
            <xsl:variable name="repository_identifier">
                <xsl:choose>
                    <xsl:when test="contains($institution,$repository)"/>
                    <xsl:when test="$institution"><xsl:value-of select="concat(replace($institution,';',' '),', ')"/></xsl:when>
                </xsl:choose>
                <xsl:value-of select="$repository"/>
            </xsl:variable>
            <xsl:variable name="title" select="//tei:msContents/tei:msItem[1]/tei:title"/>
            <xsl:variable name="quote">"</xsl:variable>
            <xsl:variable name="callNo" select="//tei:msIdentifier/tei:idno[@type='call-number']"/>
            <xsl:variable name="callNo_with_underscore" select="replace($callNo,' ','_')"/>
            <xsl:variable name="lang" select="//tei:textLang"/>
            <xsl:variable name="date" select="replace(//tei:origin/tei:p,'&quot;','')"/>
            <xsl:variable name="summary" select="replace(//tei:summary,'&quot;','')"/>
            <xsl:variable name="compuDate">
                <xsl:for-each select="//tei:origin">
                    <xsl:choose>
                        <xsl:when test="tei:origDate[1]/@notBefore"><xsl:value-of select="tei:origDate[1]/@notBefore"/>-<xsl:value-of select="tei:origDate[1]/@notAfter"/></xsl:when>
                        <xsl:when test="tei:origDate[2]/@notBefore"><xsl:value-of select="tei:origDate[2]/@notBefore"/>-<xsl:value-of select="tei:origDate[2]/@notAfter"/></xsl:when>
                        <xsl:otherwise><xsl:value-of select="tei:origDate[1]/@when"/></xsl:otherwise>
                    </xsl:choose>
                </xsl:for-each>
            </xsl:variable>
            <xsl:variable name="place" select="//tei:origin/tei:origPlace"></xsl:variable>
            <xsl:variable name="keywords"><xsl:for-each select="//tei:keywords[@n='keywords']/tei:term"><xsl:value-of select="."/><xsl:if test="following-sibling::tei:term">;</xsl:if></xsl:for-each></xsl:variable>
            <xsl:variable name="msCode" select="tokenize(replace(document-uri(.),'_TEI.xml',''),'/') [position() = last()]"/>
            <xsl:variable name="repCode">
                <xsl:choose>
                    <xsl:when test="contains($repository,'Free')">0023</xsl:when>
                    <xsl:when test="contains($repository,'Othmer')">0025</xsl:when>
                    <xsl:when test="contains($repository,'Physicians')">0027</xsl:when>
                    <xsl:when test="contains($repository,'McCabe')">0008</xsl:when>
                </xsl:choose>
            </xsl:variable>
            <xsl:variable name="path">http://openn.library.upenn.edu/Data/<xsl:value-of select="$repCode"/>/</xsl:variable>
            <xsl:variable name="cover" select="//tei:surface[1]/tei:graphic[contains(@url,'web')]/@url"/>
            <xsl:variable name="opennNav" select="concat($path,'html/',$msCode,'.html')"/>
            <xsl:variable name="opennData" select="concat($path,$msCode,'/data/')"/>
            <xsl:variable name="pageTurn" select="concat('http://138.197.87.173/bibliophilly/BookReaders/',$msCode,'/index.html#page/1/mode/2up')"/>
            <xsl:variable name="collation">
                <xsl:choose>
                    <xsl:when test="starts-with(//tei:collation/tei:p[1],'1')"><xsl:value-of select="concat('http://138.197.87.173/bibliophilly/collation/',$msCode,'/',$msCode,'.html')"/></xsl:when>
                    <xsl:otherwise/>
                </xsl:choose>
            </xsl:variable>
            <xsl:variable name="record" select="//tei:altIdentifier[@type='resource']/tei:idno"/>
"<xsl:value-of select="normalize-space($repository_identifier)"></xsl:value-of>","<xsl:value-of select="translate($title,$quote,concat($quote,$quote))"/>","<xsl:value-of select="$callNo"/>","<xsl:if test="not($lang)">None noted</xsl:if><xsl:value-of select="$lang"/>","<xsl:if test="not($date)">None noted</xsl:if><xsl:value-of select="$date"/>","<xsl:if test="not($compuDate)"></xsl:if><xsl:value-of select="normalize-space(translate($compuDate,'ca.?afterirbo',''))"/>","<xsl:if test="not($place)">None noted</xsl:if><xsl:value-of select="$place"/>","<xsl:value-of select="$keywords"/>","<xsl:value-of select="concat($path,$msCode,'/','data','/',$cover)"></xsl:value-of>",OPenn Record: <a href="{$opennNav}"><xsl:value-of select="$opennNav"/></a>,OPenn Data: <a href="{$opennData}"><xsl:value-of select="$opennData"/></a>,Page Turning: <a href="{$pageTurn}"><xsl:value-of select="$pageTurn"/></a>,"","<xsl:value-of select="$summary"></xsl:value-of>",Collation View: <a href="{$collation}"><xsl:value-of select="$collation"/></a>,"Online? Yes",""</xsl:for-each><xsl:for-each select="document('../viewshare/bibliophilly-ms-in-projects.xml')">
        <xsl:for-each select="//ms:Row"><xsl:variable name="repository" select="ms:Cell[1]/ms:Data"/>
    <xsl:variable name="title" select="ms:Cell[3]/ms:Data"/>
    <xsl:variable name="callNo" select="ms:Cell[2]/ms:Data"/>
    <xsl:variable name="date" select="ms:Cell[4]/ms:Data"/>
    <xsl:variable name="place" select="ms:Cell[5]/ms:Data"/>
    <xsl:variable name="lang" select="ms:Cell[6]/ms:Data"/>
    <xsl:variable name="month-photo" select="replace(ms:Cell[7]/ms:Data,'-01T00:00:00.000','')"/>
    <xsl:variable name="opennNav" select="ms:Cell[9]/ms:Data"/>
    <xsl:choose>
<xsl:when test="starts-with($opennNav,'h')"/>
<xsl:otherwise>
"<xsl:value-of select="normalize-space($repository)"></xsl:value-of>","<xsl:value-of select="$title"/>","<xsl:value-of select="$callNo"/>","<xsl:if test="not($lang)">None noted</xsl:if><xsl:value-of select="$lang"/>","<xsl:if test="not($date)">None noted</xsl:if><xsl:value-of select="$date"/>","","<xsl:if test="not($place)">None noted</xsl:if><xsl:value-of select="$place"/>","","","","","","","","","Online? No","Photographed: <xsl:value-of select="$month-photo"/>"</xsl:otherwise>  
    </xsl:choose></xsl:for-each>
</xsl:for-each></xsl:result-document></xsl:template>
    
</xsl:stylesheet>