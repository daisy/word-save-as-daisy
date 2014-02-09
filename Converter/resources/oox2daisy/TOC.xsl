<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:v="urn:schemas-microsoft-com:vml" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
	xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:myObj="urn:Daisy" exclude-result-prefixes="w pic wp dcterms xsi cp dc a r v dcmitype myObj xsl m o">
  <!--Imports all the XSLT-->
  <!--<xsl:import href ="Common1.xsl"/>
	<xsl:import href ="Common2.xsl"/>
	<xsl:import href ="Common3.xsl"/>
	<xsl:import href ="OOML2MML.xsl"/>-->
  <!--Parameter citation-->
  <xsl:param name="Cite_style" select="myObj:Citation()"/>
  <xsl:output method="xml" indent="no" />

  <!--template for frontmatter elements-->
  <xsl:template name="TableOfContents">
    <xsl:param name="custom"/>
    <xsl:message terminate="no">progress:frontmatter</xsl:message>

      <!--checking for Table of content Element-->
      <xsl:for-each select="document('word/document.xml')//w:document/w:body/node()">
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <!--Checking for w:p element-->
        <xsl:if test="name()='w:p'">
          <!-- Checking for TOC style-->
          <xsl:if test="w:pPr/w:pStyle[substring(@w:val,1,3)='TOC']">
            <!--Checking for w:hyperlink Element-->
            <xsl:if test="w:hyperlink">
              <xsl:variable name="aquote">"</xsl:variable>
              <xsl:variable name="set_list" select="myObj:Set_Toc()"/>
              <xsl:if test="$set_list=1">
                <!--Opening level1 and list Tag-->
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','level1','&gt;')"/>
                <xsl:if test="w:pPr/w:pStyle[@w:val='TOCHeading']">
                  <h1>
                    <xsl:value-of select="w:r/w:t"/>
                  </h1>
                </xsl:if>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','list ','type=',$aquote,'pl',$aquote,'&gt;')"/>
              </xsl:if>
              <!--Checking for w:hyperlink Element-->
              <xsl:if test="w:hyperlink and not(w:pPr/w:pStyle[@w:val='TOCHeading'])">
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','li','&gt;')"/>
                <xsl:for-each select="w:hyperlink/w:r/w:rPr/w:rStyle[@w:val='Hyperlink']">
                  <xsl:message terminate="no">progress:parahandler</xsl:message>
                  <xsl:variable name="club">
                    <xsl:value-of select="../../w:t"/>
                  </xsl:variable>
                  <xsl:value-of select="myObj:SetTOCMessage($club)"/>
                </xsl:for-each>
                <lic>
                  <xsl:value-of select="myObj:GetTOCMessage()"/>
                </lic>
                <xsl:value-of select="myObj:NullMsg()"/>
                <xsl:for-each select="w:hyperlink/w:r">
                  <xsl:message terminate="no">progress:parahandler</xsl:message>
                  <xsl:if test="not(w:rPr/w:rStyle[@w:val='Hyperlink'])and w:t">
                    <lic>
                      <xsl:attribute name="class">pagenum</xsl:attribute>
                      <xsl:text>  </xsl:text>
                      <xsl:value-of select="w:t"/>
                    </lic>
                  </xsl:if>
                </xsl:for-each>
                <!--Closing li Tag-->
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/li','&gt;')"/>
              </xsl:if>
            </xsl:if>
          </xsl:if>
        </xsl:if>
      </xsl:for-each>
      <xsl:if test="myObj:Set_Toc()&gt;1">
        <!--Closing list tag-->
        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/list','&gt;')"/>
        <!--Closing level1-->
        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/level1','&gt;')"/>
      </xsl:if>
      <!-- Calling function to Reset the counter value -->
      <xsl:variable name="setlist" select="myObj:Get_Toc()"/>

      <!--checking for Table of content-->
      <xsl:for-each select="document('word/document.xml')//w:document/w:body/node()">
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <!--Checking for w:p Tag-->
        <xsl:if test="name()='w:p'">
          <!--Checking for TOC Style-->
          <xsl:if test="w:pPr/w:pStyle[substring(@w:val,1,3)='TOC']">
            <xsl:if test ="not(w:hyperlink)">
              <xsl:variable name="aquote">"</xsl:variable>
              <xsl:variable name="set_list" select="myObj:Set_Toc()"/>
              <xsl:if test="$set_list=1">
                <!--opening level1 and list Tag-->
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','level1','&gt;')"/>
                <xsl:if test="w:pPr/w:pStyle[@w:val='TOCHeading']">
                  <h1>
                    <xsl:value-of select="w:r/w:t"/>
                  </h1>
                </xsl:if>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','list ','type=',$aquote,'pl',$aquote,'&gt;')"/>
              </xsl:if>
              <xsl:for-each select=".">
                <xsl:message terminate="no">progress:parahandler</xsl:message>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','li','&gt;')"/>
                <xsl:for-each select="w:r">
                  <xsl:message terminate="no">progress:parahandler</xsl:message>
                  <xsl:if test="w:t">
                    <xsl:variable name="setToc" select="myObj:Set_tabToc()"/>
                    <xsl:choose>
                      <xsl:when test="$setToc&gt;=2">
                        <lic>
                          <xsl:attribute name="class">
                            <xsl:value-of select="'pagenum'"/>
                          </xsl:attribute>
                          <xsl:text> </xsl:text>
                          <xsl:value-of select="w:t"/>
                        </lic>
                      </xsl:when>
                      <xsl:otherwise>
                        <lic>
                          <xsl:value-of select="w:t"/>
                        </lic>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>
                </xsl:for-each>
                <xsl:variable name="setToc" select="myObj:Get_tabToc()"/>
                <!--Closing the li Tag-->
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/li','&gt;')"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:if>
        </xsl:if>
      </xsl:for-each>
      <xsl:if test="myObj:Set_Toc()&gt;1">
        <!--Closing list Tag-->
        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/list','&gt;')"/>
        <!--Closing level1 Tag-->
        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/level1','&gt;')"/>
      </xsl:if>
      <!--Calling function which resets the counter value for TOC-->
      <xsl:variable name="set_li" select="myObj:Get_Toc()"/>

      <!--Checking for Table of content-->
      <xsl:for-each select="document('word/document.xml')//w:body/w:sdt">
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:variable name="aquote">"</xsl:variable>
        <!--Checking for Table of content-->
        <xsl:if test="w:sdtPr/w:docPartObj/w:docPartGallery/@w:val='Table of Contents'">
          <level1>
            <!--Creating class attribute-->
            <xsl:attribute name="class">
              <xsl:value-of select="'print_toc'"/>
            </xsl:attribute>
            <xsl:variable name="occurToc" select="myObj:CheckTocOccur()"/>
            <xsl:if test="$custom='Automatic'">
              <!--Calling countpageTOC template to check number of pages before TOC-->
              <!--<xsl:call-template name="countpageTOC"/>-->
            </xsl:if>
            <h1>
              <xsl:value-of select="w:sdtContent/w:p/w:r/w:t"/>
            </h1>
            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','list ','type=',$aquote,'pl',$aquote,'&gt;')"/>
            <!-- if Automatic TOC -->
            <xsl:if test="w:sdtContent/w:p/w:hyperlink">
              <xsl:for-each select="w:sdtContent/w:p/w:hyperlink">
                <xsl:message terminate="no">progress:parahandler</xsl:message>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','li','&gt;')"/>
                <xsl:for-each select="w:r/w:rPr/w:rStyle[@w:val='Hyperlink']">
                  <xsl:message terminate="no">progress:parahandler</xsl:message>
                  <xsl:variable name="club">
                    <xsl:value-of select="../../w:t"/>
                  </xsl:variable>
                  <xsl:value-of select="myObj:SetTOCMessage($club)"/>
                </xsl:for-each>
                <lic>
                  <xsl:value-of select="myObj:GetTOCMessage()"/>
                </lic>
                <xsl:value-of select="myObj:NullMsg()"/>
                <xsl:for-each select="w:r">
                  <xsl:message terminate="no">progress:parahandler</xsl:message>
                  <xsl:if test="not(w:rPr/w:rStyle[@w:val='Hyperlink']) and w:t">
                    <lic>
                      <xsl:attribute name="class">pagenum</xsl:attribute>
                      <xsl:text>  </xsl:text>
                      <xsl:value-of select="w:t"/>
                    </lic>
                  </xsl:if>
                </xsl:for-each>
                <!--Closing li Tag-->
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/li','&gt;')"/>
              </xsl:for-each>
            </xsl:if>
            <!-- if Manual TOC -->
            <xsl:if test="not(w:sdtContent/w:p/w:hyperlink)">
              <xsl:for-each select="w:sdtContent/w:p">
                <xsl:if test="not(w:pPr/w:pStyle[@w:val='TOCHeading'])">
                  <xsl:message terminate="no">progress:parahandler</xsl:message>
                  <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','li','&gt;')"/>
                  <xsl:for-each select="w:r/w:t">
                    <xsl:message terminate="no">progress:parahandler</xsl:message>
                    <xsl:variable name="club">
                      <xsl:value-of select="."/>
                    </xsl:variable>
                    <xsl:value-of select="myObj:SetTOCMessage($club)"/>
                  </xsl:for-each>
                  <lic>
                    <xsl:value-of select="myObj:GetTOCMessage()"/>
                  </lic>
                  <xsl:value-of select="myObj:NullMsg()"/>
                  <!--Closing li Tag-->
                  <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/li','&gt;')"/>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
            <xsl:if test="(not(following-sibling::node()[1][w:r/w:rPr/w:rStyle[substring(@w:val,1,15)='PageNumberDAISY']]) and ($custom='Custom')) or (not($custom='Custom'))">
              <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/list','&gt;')"/>
            </xsl:if>
          </level1>
        </xsl:if>
      </xsl:for-each>
    
  </xsl:template>

</xsl:stylesheet>
  