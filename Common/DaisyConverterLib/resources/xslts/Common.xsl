<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="2.0"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                xmlns:dcterms="http://purl.org/dc/terms/"
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                xmlns:dc="http://purl.org/dc/elements/1.1/"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:v="urn:schemas-microsoft-com:vml"
                xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                xmlns:o="urn:schemas-microsoft-com:office:office"
                xmlns:d="DaisyClass"
                xmlns="http://www.daisy.org/z3986/2005/dtbook/"
                exclude-result-prefixes="w pic wp dcterms xsi cp dc a r v dcmitype d xsl m o">
    <!--Parameter citation-->
    <xsl:param name="Cite_style" select="d:Citation($myObj)"/>

    <!--template for frontmatter elements-->
    <xsl:template name="FrontMatter">
        <!--Parameter Tile of the document-->
        <xsl:param name="Title"></xsl:param>
        <!--Parameter author of the document-->
        <xsl:param name="Creator"></xsl:param>
        <!--Parameter trackchanges-->
        <xsl:param name="prmTrack"/>
        <!--Parameter version of Office-->
        <xsl:param name="version"/>
        <!--Parameter custom page number-->
        <xsl:param name="custom"/>
        <xsl:param name="masterSub"/>
        <xsl:param name="sOperators"/>
        <xsl:param name="sMinuses"/>
        <xsl:param name="sNumbers"/>
        <xsl:param name="sZeros"/>
        <xsl:param name="imgOption"/>
        <xsl:param name="dpi"/>
        <xsl:param name="charStyles"/>

        <xsl:message terminate="no">progress:frontmatter</xsl:message>

        <doctitle>
            <xsl:choose>
                <!--Taking Document Title value from core.xml-->
                <xsl:when test="string-length($Title) = 0">
                    <xsl:value-of select="$docPropsCoreXml//cp:coreProperties/dc:title"/>
                </xsl:when>
                <!--Taking the Title value entered by the user-->
                <xsl:otherwise>
                    <xsl:value-of select="$Title"/>
                </xsl:otherwise>
            </xsl:choose>
        </doctitle>
        <docauthor>
            <xsl:choose>
                <!--Taking Document creator value from core.xml-->
                <xsl:when test="string-length($Creator) = 0">
                    <xsl:value-of select="$docPropsCoreXml//cp:coreProperties/dc:creator"/>
                </xsl:when>
                <!--Taking the Creator value entered by the user-->
                <xsl:otherwise>
                    <xsl:value-of select="$Creator"/>
                </xsl:otherwise>
            </xsl:choose>
        </docauthor>

        <xsl:call-template name="Matter">
            <xsl:with-param name="prmTrack" select="$prmTrack"/>
            <xsl:with-param name="version" select="$version"/>
            <xsl:with-param name="custom" select="$custom"/>
            <xsl:with-param name="masterSub" select="$masterSub"/>
            <xsl:with-param name="sOperators" select="$sOperators"/>
            <xsl:with-param name="sMinuses" select="$sMinuses"/>
            <xsl:with-param name="sNumbers" select="$sNumbers"/>
            <xsl:with-param name="sZeros" select="$sZeros"/>
            <xsl:with-param name="imgOption" select="$imgOption"/>
            <xsl:with-param name="dpi" select="$dpi"/>
            <xsl:with-param name="charStyles" select="$charStyles"/>
            <xsl:with-param name="matterType" select="'Frontmatter'" />
        </xsl:call-template>

    </xsl:template>


    <!--Template for implementing bodymatter elements-->
    <xsl:template name="BodyMatter">
        <!--Parameter trackchanges-->
        <xsl:param name="prmTrack"/>
        <!--Parameter version of Office-->
        <xsl:param name="version"/>
        <!--Parameter custom page number-->
        <xsl:param name="custom"/>
        <xsl:param name="masterSub"/>
        <xsl:param name="sOperators"/>
        <xsl:param name="sMinuses"/>
        <xsl:param name="sNumbers"/>
        <xsl:param name="sZeros"/>
        <xsl:param name="imgOption"/>
        <xsl:param name="dpi"/>
        <xsl:param name="charStyles"/>

        <xsl:call-template name="Matter">
            <xsl:with-param name="prmTrack" select="$prmTrack"/>
            <xsl:with-param name="version" select="$version"/>
            <xsl:with-param name="custom" select="$custom"/>
            <xsl:with-param name="masterSub" select="$masterSub"/>
            <xsl:with-param name="sOperators" select="$sOperators"/>
            <xsl:with-param name="sMinuses" select="$sMinuses"/>
            <xsl:with-param name="sNumbers" select="$sNumbers"/>
            <xsl:with-param name="sZeros" select="$sZeros"/>
            <xsl:with-param name="imgOption" select="$imgOption"/>
            <xsl:with-param name="dpi" select="$dpi"/>
            <xsl:with-param name="charStyles" select="$charStyles"/>
            <xsl:with-param name="matterType" select="'Bodymatter'" />
        </xsl:call-template>

    </xsl:template>



    <!--Template for implementing Rearmatter elements-->
    <xsl:template name="RearMatter">
        <!--Parameter trackchanges-->
        <xsl:param name="prmTrack"/>
        <!--Parameter version of Office-->
        <xsl:param name="version"/>
        <!--Parameter custom page number-->
        <xsl:param name="custom"/>
        <xsl:param name="masterSub"/>
        <xsl:param name="sOperators"/>
        <xsl:param name="sMinuses"/>
        <xsl:param name="sNumbers"/>
        <xsl:param name="sZeros"/>
        <xsl:param name="imgOption"/>
        <xsl:param name="dpi"/>
        <xsl:param name="charStyles"/>

        <xsl:if test="not(count($documentXml//w:document/w:body/w:p/w:r/w:rPr/w:rStyle[@w:val='EndnoteReference'])=0)">
            <level1>
                <!--Checking if any elements should be translated to the rearmatter-->
                <!--Otherwise Traversing through document.xml file and passing the Endnote id to the Note template.-->
                <xsl:for-each select="$documentXml//w:document/w:body/w:p/w:r/w:rPr/w:rStyle">
                    <xsl:message terminate="no">progress:parahandler</xsl:message>
                    <xsl:if test="@w:val='EndnoteReference'">
                        <xsl:call-template name="tmpNote">
                            <xsl:with-param name="endNoteId" select="../../w:endnoteReference/@w:id"/>
                            <xsl:with-param name="sOperators" select="$sOperators"/>
                            <xsl:with-param name="sMinuses" select="$sMinuses"/>
                            <xsl:with-param name="sNumbers" select="$sNumbers"/>
                            <xsl:with-param name="sZeros" select="$sZeros"/>
                            <xsl:with-param name="vernote" select="$version"/>
                        </xsl:call-template>
                    </xsl:if>
                </xsl:for-each>
            </level1>
        </xsl:if>

        <xsl:call-template name="Matter">
            <xsl:with-param name="prmTrack" select="$prmTrack"/>
            <xsl:with-param name="version" select="$version"/>
            <xsl:with-param name="custom" select="$custom"/>
            <xsl:with-param name="masterSub" select="$masterSub"/>
            <xsl:with-param name="sOperators" select="$sOperators"/>
            <xsl:with-param name="sMinuses" select="$sMinuses"/>
            <xsl:with-param name="sNumbers" select="$sNumbers"/>
            <xsl:with-param name="sZeros" select="$sZeros"/>
            <xsl:with-param name="imgOption" select="$imgOption"/>
            <xsl:with-param name="dpi" select="$dpi"/>
            <xsl:with-param name="charStyles" select="$charStyles"/>
            <xsl:with-param name="matterType" select="'Rearmatter'" />
        </xsl:call-template>

    </xsl:template>



    <xsl:template name="Matter">
        <!--Parameter trackchanges-->
        <xsl:param name="prmTrack"/>
        <!--Parameter version of Office-->
        <xsl:param name="version"/>
        <!--Parameter custom page number-->
        <xsl:param name="custom"/>
        <xsl:param name="masterSub"/>
        <xsl:param name="sOperators"/>
        <xsl:param name="sMinuses"/>
        <xsl:param name="sNumbers"/>
        <xsl:param name="sZeros"/>
        <xsl:param name="imgOption"/>
        <xsl:param name="dpi"/>
        <xsl:param name="charStyles"/>
        <xsl:param name="matterType" />
        <!--Variable external images-->
        <xsl:variable name="external">
            <!--Calling d:ExternalImage() to check external images-->
            <xsl:value-of select="d:ExternalImage($myObj)"/>
        </xsl:variable>
        <!--Fedility loss External images-->
        <xsl:if test="$external='translation.oox2Daisy.ExternalImage'">
            <xsl:message terminate="no">translation.oox2Daisy.ExternalImage</xsl:message>
        </xsl:if>

        <xsl:variable name="var_heading">
            <xsl:for-each select="$stylesXml//w:styles/w:style/w:name[@w:val='heading 1']">
                <xsl:message terminate="no">progress:parahandler</xsl:message>
                <xsl:value-of select="../@w:styleId"/>
            </xsl:for-each>
        </xsl:variable>
        <!--Looping through each hyperlink-->
        <xsl:for-each select="$documentXml//w:document/w:body/w:p/w:hyperlink">
            <xsl:message terminate="no">progress:parahandler</xsl:message>
            <!--Calling d:AddHyperlink() for storing Anchor in Hyperlink-->
            <xsl:sequence select="d:sink(d:AddHyperlink($myObj,string(@w:anchor)))"/> <!-- empty -->
        </xsl:for-each>
        <xsl:message terminate="no">progress:parahandler</xsl:message>


        <!--<xsl:if test="$matterType=''">-->
        <!--Checking the first paragraph of the document-->
        <xsl:for-each select="$documentXml//w:document/w:body/w:p[1]">
            <xsl:message terminate="no">progress:parahandler</xsl:message>
            <!--If first paragraph is not Heading-->
            <xsl:if test="not(w:pPr/w:pStyle[substring(@w:val,1,7)='Heading']) or $matterType='Rearmatter'">
                <!--Calling Template to add level1-->
                <xsl:call-template name="AddLevel">
                    <xsl:with-param name="levelValue" select="1"/>
                    <xsl:with-param name="check" select="0"/>
                    <xsl:with-param name="verhead" select="$version"/>
                    <xsl:with-param name="custom" select="$custom"/>
                    <xsl:with-param name="sOperators" select="$sOperators"/>
                    <xsl:with-param name="sMinuses" select="$sMinuses"/>
                    <xsl:with-param name="sNumbers" select="$sNumbers"/>
                    <xsl:with-param name="sZeros" select="$sZeros"/>
                    <xsl:with-param name="mastersubhead" select="$masterSub"/>
                    <xsl:with-param name="txt" select="'0'"/>
                    <xsl:with-param name="lvlcharStyle" select="$charStyles"/>
                </xsl:call-template>
            </xsl:if>
        </xsl:for-each>
        <!--</xsl:if>-->

        <xsl:sequence select="d:ResetCurrentMatterType($myObj)"/> <!-- empty -->
        <!--Traversing through each node of the document-->
        <xsl:for-each select="$documentXml//w:body/node()">
            <xsl:message terminate="no">progress:parahandler</xsl:message>
            <xsl:if test="not($masterSub='Yes')">
                <xsl:call-template name="SetCurrentMatterType" />
            </xsl:if>
            <xsl:if test="d:GetCurrentMatterType($myObj)=$matterType">
                <xsl:choose>
                    <!--Checking for Praragraph element-->
                    <xsl:when test="name()='w:p'
                                                                                                and not(*/w:pStyle[substring(@w:val, 1, 11)='Frontmatter'])
                                                                                                and not(*/w:pStyle[substring(@w:val, 1, 10)='Bodymatter'])
                                                                                                and not(*/w:pStyle[substring(@w:val, 1, 10)='Rearmatter'])">
                        <xsl:sequence select="d:sink(d:CheckCaverPage($myObj))"/> <!-- empty -->
                        <xsl:call-template name="StyleContainer">
                            <xsl:with-param name="prmTrack" select="$prmTrack"/>
                            <xsl:with-param name="VERSION" select="$version"/>
                            <xsl:with-param name="custom" select="$custom"/>
                            <xsl:with-param name="styleHeading" select="$var_heading"/>
                            <xsl:with-param name="mastersubstyle" select="$masterSub"/>
                            <xsl:with-param name="imgOptionStyle" select="$imgOption"/>
                            <xsl:with-param name="dpiStyle" select="$dpi"/>
                            <xsl:with-param name="characterStyle" select="$charStyles"/>
                        </xsl:call-template>
                    </xsl:when>
                    <!--Checking for Table element-->
                    <xsl:when test="name()='w:tbl'">
                        <!--If exists then calling TableHandler-->
                        <xsl:sequence select="d:sink(d:CheckCaverPage($myObj))"/> <!-- empty -->
                        <xsl:call-template name="TableHandler">
                            <xsl:with-param name="parmVerTable" select="$version"/>
                            <xsl:with-param name="custom" select="$custom"/>
                            <xsl:with-param name="mastersubtbl" select="$masterSub"/>
                            <xsl:with-param name="characterStyle" select="$charStyles"/>
                        </xsl:call-template>
                    </xsl:when>
                    <!--Checking for section element-->
                    <xsl:when test="name()='w:sectPr'">
                        <!--calling for foonote template and displaying footnote text at the end of the page-->
                        <xsl:call-template name="footnote">
                            <xsl:with-param name="verfoot" select="$version"/>
                            <xsl:with-param name="mastersubfoot" select="$masterSub"/>
                            <xsl:with-param name="sOperators" select="$sOperators"/>
                            <xsl:with-param name="sMinuses" select="$sMinuses"/>
                            <xsl:with-param name="sNumbers" select="$sNumbers"/>
                            <xsl:with-param name="sZeros" select="$sZeros"/>
                        </xsl:call-template>

                        <xsl:if test="d:CheckCaverPage($myObj)=1">
                            <level1>
                                <p/>
                            </level1>
                        </xsl:if>
                    </xsl:when>
                    <!--Checking for Structured document element-->
                    <xsl:when test="name()='w:sdt'">
                        <xsl:choose>
                            <xsl:when test="w:sdtPr/w:docPartObj/w:docPartGallery/@w:val='Table of Contents'">
                                <!--Save Level before closing all levels-->
                                <xsl:variable name="PeekLevel" select="d:PeekLevel($myObj)"/>
                                <!--Close all levels before Table Of Contents-->
                                <xsl:call-template name="CloseLevel">
                                    <xsl:with-param name="CurrentLevel" select="1"/>
                                </xsl:call-template>
                                <!--Calling Template to add Table Of Contents-->
                                <xsl:call-template name="TableOfContents">
                                    <xsl:with-param name="custom" select="$custom"/>
                                </xsl:call-template>
                                <!--Open $PeekLevel levels after Table Of Contents-->
                                <xsl:call-template name="AddLevel">
                                    <xsl:with-param name="levelValue" select="$PeekLevel"/>
                                    <xsl:with-param name="check" select="1"/>
                                    <xsl:with-param name="verhead" select="$version"/>
                                    <xsl:with-param name="custom" select="$custom"/>
                                    <xsl:with-param name="sOperators" select="$sOperators"/>
                                    <xsl:with-param name="sMinuses" select="$sMinuses"/>
                                    <xsl:with-param name="sNumbers" select="$sNumbers"/>
                                    <xsl:with-param name="sZeros" select="$sZeros"/>
                                    <xsl:with-param name="mastersubhead" select="$masterSub"/>
                                    <xsl:with-param name="txt" select="0"/>
                                    <xsl:with-param name="lvlcharStyle" select="$charStyles"/>
                                </xsl:call-template>
                            </xsl:when>
                            <xsl:otherwise>
                                <!--Calling structuredocument for handling Fidelity loss-->
                                <xsl:call-template name="structuredocument"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:when>
                    <!--Checking for BookmarkStart element-->
                    <xsl:when test="name()='w:bookmarkStart'">
                        <!--Checking Whether BookMarkStart is related to Abbreviations or not -->
                        <xsl:if test="substring(@w:name,1,13)='Abbreviations'">
                            <!--Storing the full form of Abbreviation in a variable-->
                            <xsl:variable name="full">
                                <xsl:value-of select="d:FullAbbr($myObj,@w:name,$version)"/>
                            </xsl:variable>
                            <xsl:choose>
                                <!--Checking whether all previous Abbreviations tags are closed or not before opening an new Abbreviation tag-->
                                <xsl:when test ="not(d:AbbrAcrFlag($myObj)=1)">
                                    <xsl:choose>
                                        <!--checking whether an Abbreviation is having Full Form or not-->
                                        <xsl:when test="not($full='')">
                                            <xsl:variable name="aquote">"</xsl:variable>
                                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','abbr ','title=',$aquote,$full,$aquote,'&gt;')"/>
                                            <xsl:sequence select="d:sink(d:SetAbbrAcrFlag($myObj))"/> <!-- empty -->
                                        </xsl:when>
                                        <xsl:otherwise>
                                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','abbr','&gt;')"/>
                                            <xsl:sequence select="d:sink(d:SetAbbrAcrFlag($myObj))"/> <!-- empty -->
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:choose>
                                        <!--checking whether an Abbreviation is having Full Form or not-->
                                        <xsl:when test="not($full='')">
                                            <xsl:variable name="aquote">"</xsl:variable>
                                            <xsl:variable name="temp" select="concat('&lt;','abbr ','title=',$aquote,$full,$aquote,'&gt;')"/>
                                            <xsl:sequence select="d:sink(d:PushAbrAcr($myObj,$temp))"/> <!-- empty -->
                                        </xsl:when>
                                        <xsl:otherwise>
                                            <xsl:variable name="aquote">"</xsl:variable>
                                            <xsl:variable name="temp" select="concat('&lt;','abbr','&gt;')"/>
                                            <xsl:sequence select="d:sink(d:PushAbrAcr($myObj,$temp))"/> <!-- empty -->
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:if>
                        <!--Checking Whether BookMarkStart is related to Acronyms or not -->
                        <xsl:if test="substring(@w:name,1,11)='AcronymsYes'">
                            <!--Storing the full form of Abbreviation in a variable-->
                            <xsl:variable name="full">
                                <xsl:value-of select="d:FullAcr($myObj,@w:name,$version)"/>
                            </xsl:variable>
                            <xsl:choose>
                                <!--Checking whether all previous Acronyms tags are closed or not before opening an new Acronyms tag-->
                                <xsl:when test ="not(d:AbbrAcrFlag($myObj)=1)">
                                    <xsl:choose>
                                        <!--checking whether an Abbreviation is having Full Form or not-->
                                        <xsl:when test="not($full='')">
                                            <xsl:variable name="aquote">"</xsl:variable>
                                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','acronym ','pronounce=',$aquote,'yes',$aquote,' title=',$aquote,$full,$aquote,'&gt;')"/>
                                            <xsl:sequence select="d:sink(d:SetAbbrAcrFlag($myObj))"/> <!-- empty -->
                                        </xsl:when>
                                        <xsl:otherwise>
                                            <xsl:variable name="aquote">"</xsl:variable>
                                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','acronym ','pronounce=',$aquote,'yes',$aquote,'&gt;')"/>
                                            <xsl:sequence select="d:sink(d:SetAbbrAcrFlag($myObj))"/> <!-- empty -->
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:choose>
                                        <!--checking whether an Abbreviation is having Full Form or not-->
                                        <xsl:when test="not($full='')">
                                            <xsl:variable name="aquote">"</xsl:variable>
                                            <xsl:variable name="temp" select="concat('&lt;','acronym ','pronounce=',$aquote,'yes',$aquote,' title=',$aquote,$full,$aquote,'&gt;')"/>
                                            <xsl:sequence select="d:sink(d:PushAbrAcr($myObj,$temp))"/> <!-- empty -->
                                        </xsl:when>
                                        <xsl:otherwise>
                                            <xsl:variable name="aquote">"</xsl:variable>
                                            <xsl:variable name="temp" select="concat('&lt;','acronym ','pronounce=',$aquote,'yes',$aquote,'&gt;')"/>
                                            <xsl:sequence select="d:sink(d:PushAbrAcr($myObj,$temp))"/> <!-- empty -->
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:if>
                        <!--Checking Whether BookMarkStart is related to Acronyms or not -->
                        <xsl:if test="substring(@w:name,1,10)='AcronymsNo'">
                            <!--Storing the full form of Abbreviation in a variable-->
                            <xsl:variable name="full">
                                <xsl:value-of select="d:FullAcr($myObj,@w:name,$version)"/>
                            </xsl:variable>
                            <xsl:choose>
                                <!--Checking whether all previous Acronyms tags are closed or not before opening an new Acronyms tag-->
                                <xsl:when test ="not(d:AbbrAcrFlag($myObj)=1)">
                                    <xsl:choose>
                                        <!--checking whether an Abbreviation is having Full Form or not-->
                                        <xsl:when test="not($full='')">
                                            <xsl:variable name="aquote">"</xsl:variable>
                                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','acronym ','pronounce=',$aquote,'no',$aquote,' title=',$aquote,$full,$aquote,'&gt;')"/>
                                            <xsl:sequence select="d:sink(d:SetAbbrAcrFlag($myObj))"/> <!-- empty -->
                                        </xsl:when>
                                        <xsl:otherwise>
                                            <xsl:variable name="aquote">"</xsl:variable>
                                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','acronym ','pronounce=',$aquote,'no',$aquote,'&gt;')"/>
                                            <xsl:sequence select="d:sink(d:SetAbbrAcrFlag($myObj))"/> <!-- empty -->
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:choose>
                                        <!--checking whether an Abbreviation is having Full Form or not-->
                                        <xsl:when test="not($full='')">
                                            <xsl:variable name="aquote">"</xsl:variable>
                                            <xsl:variable name="temp" select="concat('&lt;','acronym ','pronounce=',$aquote,'no',$aquote,' title=',$aquote,$full,$aquote,'&gt;')"/>
                                            <xsl:sequence select="d:sink(d:PushAbrAcr($myObj,$temp))"/> <!-- empty -->
                                        </xsl:when>
                                        <xsl:otherwise>
                                            <xsl:variable name="aquote">"</xsl:variable>
                                            <xsl:variable name="temp" select="concat('&lt;','acronym ','pronounce=',$aquote,'no',$aquote,'&gt;')"/>
                                            <xsl:sequence select="d:sink(d:PushAbrAcr($myObj,$temp))"/> <!-- empty -->
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:if>
                    </xsl:when>
                    <!--Checking for BookMarkEnd -->
                    <xsl:when test="name()='w:bookmarkEnd'">
                        <xsl:variable name="seperate">
                            <xsl:variable name="id" select="@w:id"/>
                            <xsl:variable name="tempAbbr" select="$documentXml//w:bookmarkStart[@w:id=$id]/@w:name"/>
                            <xsl:value-of select="d:Book($myObj,$tempAbbr)"/>
                        </xsl:variable>
                        <!--Checking whether BookMarkEnd is related to Abbreviations or not -->
                        <xsl:if test="$seperate='AbbrTrue'">
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/abbr','&gt;')"/>
                            <xsl:sequence select="d:sink(d:ReSetAbbrAcrFlag($myObj))"/> <!-- empty -->
                            <xsl:if test="d:CountAbrAcrpara($myObj) &gt; 0">
                                <xsl:value-of disable-output-escaping="yes" select="d:PeekAbrAcrpara($myObj)"/>
                            </xsl:if>
                            <xsl:if test="d:CountAbrAcrhead($myObj) &gt; 0">
                                <xsl:value-of disable-output-escaping="yes" select="d:PeekAbrAcrhead($myObj)"/>
                            </xsl:if>
                        </xsl:if>
                        <!--Checking whether BookMarkEnd is related to Acronyms or not -->
                        <xsl:if test="$seperate='AcrTrue'">
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/acronym','&gt;')"/>
                            <xsl:sequence select="d:sink(d:ReSetAbbrAcrFlag($myObj))"/> <!-- empty -->
                            <xsl:if test="d:CountAbrAcrpara($myObj) &gt; 0">
                                <xsl:value-of disable-output-escaping="yes" select="d:PeekAbrAcrpara($myObj)"/>
                            </xsl:if>
                            <xsl:if test="d:CountAbrAcrhead($myObj) &gt; 0">
                                <xsl:value-of disable-output-escaping="yes" select="d:PeekAbrAcrhead($myObj)"/>
                            </xsl:if>
                        </xsl:if>
                    </xsl:when>
                    <!--Checking for Pagebreaks and calling footnote template for displaying footnote text at the end of the page-->
                    <xsl:otherwise>
                        <!--Implementing fidelity loss for unhandled elements-->
                        <xsl:message terminate="no">
                            <xsl:value-of select="concat('translation.oox2Daisy.UncoveredElement|',name())"/>
                        </xsl:message>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:if>
        </xsl:for-each>
        <!--Call a template to Close all the levels -->
        <xsl:call-template name="CloseLevel">
            <xsl:with-param name="CurrentLevel" select="1"/>
        </xsl:call-template>
    </xsl:template>


    <!--Template to implement different paragraph styles-->
    <!--Implementing all the paragraph styles and all other feature that appears inside the paragraph-->
    <xsl:template name="ParaHandler">
        <xsl:param name="flag"/>
        <xsl:param name="prmTrack"/>
        <xsl:param name="VERSION"/>
        <xsl:param name="flagNote"/>
        <xsl:param name="checkid"/>
        <xsl:param name="sOperators"/>
        <xsl:param name="sMinuses"/>
        <xsl:param name="sNumbers"/>
        <xsl:param name="sZeros"/>
        <xsl:param name="custom"/>
        <xsl:param name="txt"/>
        <xsl:param name="level"/>
        <xsl:param name="mastersubpara"/>
        <xsl:param name="imgOptionPara"/>
        <xsl:param name="dpiPara"/>
        <xsl:param name="charparahandlerStyle"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:variable name="quote">"</xsl:variable>
        <!--Calling Footnote template when the page break is encountered.-->

        <xsl:if test="((w:r/w:lastRenderedPageBreak) or (w:r/w:br/@w:type='page')) and (not($flag='0'))">
            <!--If parent element is not table cell-->
            <xsl:if test="not(name(..)='w:tc')">
                <xsl:call-template name="footnote">
                    <xsl:with-param name="verfoot" select="$VERSION"/>
                    <xsl:with-param name="mastersubfoot" select="$mastersubpara"/>
                    <xsl:with-param name="sOperators" select="$sOperators"/>
                    <xsl:with-param name="sMinuses" select="$sMinuses"/>
                    <xsl:with-param name="sNumbers" select="$sNumbers"/>
                    <xsl:with-param name="sZeros" select="$sZeros"/>
                </xsl:call-template>
                <xsl:choose>
                    <!--Checking for page break-->
                    <xsl:when test="((w:r/w:br/@w:type='page') and not((following-sibling::w:p[1]/w:pPr/w:sectPr) or (following-sibling::w:p[2]/w:r/w:lastRenderedPageBreak) or (following-sibling::w:p[1]/w:r/w:lastRenderedPageBreak) or (following-sibling::w:sdt[1]/w:sdtPr/w:docPartObj/w:docPartGallery/@w:val='Table of Contents')))">
                        <xsl:call-template name="footnote">
                            <xsl:with-param name="verfoot" select="$VERSION"/>
                            <xsl:with-param name="mastersubfoot" select="$mastersubpara"/>
                            <xsl:with-param name="sOperators" select="$sOperators"/>
                            <xsl:with-param name="sMinuses" select="$sMinuses"/>
                            <xsl:with-param name="sNumbers" select="$sNumbers"/>
                            <xsl:with-param name="sZeros" select="$sZeros"/>
                        </xsl:call-template>
                    </xsl:when>
                    <!--Checking for page break-->
                    <xsl:when test="(w:r/w:lastRenderedPageBreak)">
                        <xsl:call-template name="footnote">
                            <xsl:with-param name="verfoot" select="$VERSION"/>
                            <xsl:with-param name="mastersubfoot" select="$mastersubpara"/>
                            <xsl:with-param name="sOperators" select="$sOperators"/>
                            <xsl:with-param name="sMinuses" select="$sMinuses"/>
                            <xsl:with-param name="sNumbers" select="$sNumbers"/>
                            <xsl:with-param name="sZeros" select="$sZeros"/>
                        </xsl:call-template>
                    </xsl:when>
                </xsl:choose>
            </xsl:if>
        </xsl:if>

        <!--Initializing section Information if page number style is automatic-->
        <xsl:if test="(w:pPr/w:sectPr) and not($flag='2') and ($custom='Automatic')">
            <xsl:call-template name="SectionInfo"/>
        </xsl:if>

        <xsl:if test="(w:r/w:pict//v:textbox/w:txbxContent) and (not(w:r/w:pict/v:group)) and ($VERSION='12.0')">
            <xsl:if test="not(w:r/w:pict//v:textbox/w:txbxContent/w:p/w:pPr/w:pStyle[@w:val='Caption'])">
                <xsl:call-template name="TempSidebar">
                    <xsl:with-param name="flag" select="$flag"/>
                    <xsl:with-param name="custom" select="$custom"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="prmTrack" select="$prmTrack"/>
                    <xsl:with-param name="VERSION" select="$VERSION"/>
                    <xsl:with-param name="charparahandlerStyle" select="$charparahandlerStyle"/>
                    <xsl:with-param name="mastersubpara" select="$mastersubpara"/>
                    <xsl:with-param name="level" select="$level"/>
                </xsl:call-template>
            </xsl:if>
        </xsl:if>

        <xsl:if test="(w:r/w:pict//v:textbox/w:txbxContent)and not(w:r/w:pict/v:group[@editas='orgchart']) and (($VERSION='11.0') or ($VERSION='10.0'))">
            <xsl:if test="not(w:r/w:pict//v:textbox/w:txbxContent/w:p/w:pPr/w:pStyle[@w:val='Caption'])">
                <xsl:call-template name="TempSidebar">
                    <xsl:with-param name="flag" select="$flag"/>
                    <xsl:with-param name="custom" select="$custom"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="prmTrack" select="$prmTrack"/>
                    <xsl:with-param name="VERSION" select="$VERSION"/>
                    <xsl:with-param name="charparahandlerStyle" select="$charparahandlerStyle"/>
                    <xsl:with-param name="mastersubpara" select="$mastersubpara"/>
                    <xsl:with-param name="level" select="$level"/>
                </xsl:call-template>
            </xsl:if>
        </xsl:if>

        <xsl:if test="not($flag='0') and not(d:AbbrAcrFlag($myObj)=1)and not($flagNote='hyper')">
            <xsl:if test="(not(d:GetTestRun($myObj)&gt;='1')) and (d:GetCodeFlag($myObj)='0')">
                <!--Setting a flag for linenumber-->
                <xsl:sequence select="d:sink(d:Setlinenumflag($myObj))"/> <!-- empty -->
                <xsl:choose>
                    <xsl:when test="(w:r/w:rPr/w:lang) or (w:r/w:rPr/w:rFonts/@w:hint)">
                        <xsl:call-template name="Languages">
                            <xsl:with-param name="Attribute" select="'0'"/>
                            <xsl:with-param name="txt" select="$txt"/>
                        </xsl:call-template>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:if test="(not(w:r/w:rPr/w:rStyle[@w:val='PageNumberDAISY']) and ($custom='Custom')) or (not($custom='Custom'))">
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p','&gt;')"/>
                        </xsl:if>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:if>

            <!--Adding Note index-->
            <xsl:if test="$flagNote='Note'">
                <xsl:if test="d:NoteFlag($myObj)=1">
                    <xsl:value-of select="$checkid - 1"/>
                </xsl:if>
            </xsl:if>
        </xsl:if>

        <!--Traversing through each node inside a paragraph-->
        <xsl:for-each select="./node()">
            <xsl:message terminate="no">progress:parahandler</xsl:message>

            <xsl:if test="name()='w:subDoc'">
                <xsl:if test="$mastersubpara='Yes'">
                    <xsl:variable name="aquote">"</xsl:variable>
                    <xsl:variable name="temp"    select="concat('&lt;','subdoc ','rId=',$aquote,@r:id,$aquote,'&gt;','&lt;','/subdoc','&gt;')"/>
                    <xsl:sequence select="d:sink(d:PushMasterSubdoc($myObj,$temp))"/> <!-- empty -->
                    <xsl:sequence select="d:sink(d:MasterSubSetFlag($myObj))"/> <!-- empty -->
                </xsl:if>
            </xsl:if>

            <!--Checking condition for MathEquations in Word2007-->
            <xsl:if test="m:oMathPara">
                <xsl:call-template name="ooml2mml">
                    <xsl:with-param name="sOperators" select="$sOperators"/>
                    <xsl:with-param name="sMinuses" select="$sMinuses"/>
                    <xsl:with-param name="sNumbers" select="$sNumbers"/>
                    <xsl:with-param name="sZeros" select="$sZeros"/>
                </xsl:call-template>
            </xsl:if>

            <!--Checking condition for MathEquations in Word2007-->
            <xsl:if test="m:oMath">
                <xsl:choose>
                    <!--Checking for BDO Element in MathEquation-->
                    <xsl:when test="../w:pPr/w:bidi">
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,'&gt;')"/>
                        <!--Calling Mathml Template for Math Equations-->
                        <xsl:call-template name="ooml2mml">
                            <xsl:with-param name="sOperators" select="$sOperators"/>
                            <xsl:with-param name="sMinuses" select="$sMinuses"/>
                            <xsl:with-param name="sNumbers" select="$sNumbers"/>
                            <xsl:with-param name="sZeros" select="$sZeros"/>
                        </xsl:call-template>
                        <!--Closing bdo Tag-->
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <!--Calling Mathml Template for Math Equations-->
                        <xsl:call-template name="ooml2mml">
                            <xsl:with-param name="sOperators" select="$sOperators"/>
                            <xsl:with-param name="sMinuses" select="$sMinuses"/>
                            <xsl:with-param name="sNumbers" select="$sNumbers"/>
                            <xsl:with-param name="sZeros" select="$sZeros"/>
                        </xsl:call-template>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:if>

            <!--Checking condition for MathEquations in Word2007-->
            <xsl:if test="../m:oMath">
                <!--Calling Mathml Template for Math Equations-->
                <xsl:call-template name="ooml2mml">
                    <xsl:with-param name="sOperators" select="$sOperators"/>
                    <xsl:with-param name="sMinuses" select="$sMinuses"/>
                    <xsl:with-param name="sNumbers" select="$sNumbers"/>
                    <xsl:with-param name="sZeros" select="$sZeros"/>
                </xsl:call-template>
            </xsl:if>

            <!--Checking for smartTag element-->
            <xsl:if test="name()='w:smartTag'">
                <xsl:call-template name="smartTag"/>
            </xsl:if>

            <!--Checking for fldSimple element-->
            <xsl:if test="name()='w:fldSimple'">
                <xsl:call-template name="fldSimple">
                    <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                </xsl:call-template>
            </xsl:if>

            <!--Checking for Hyperlink element-->
            <xsl:if test="name()='w:hyperlink' and not(preceding-sibling::w:pPr/w:pStyle[substring(@w:val,1,3)='TOC'])">
                <xsl:if test="d:GetTestRun($myObj)&gt;='1'">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/a','&gt;')"/>
                    <xsl:sequence select="d:sink(d:SetBookmark($myObj))"/> <!-- empty -->
                </xsl:if>
                <a>
                    <xsl:choose>
                        <!--If both id and anchor attribute is present in hyperlink-->
                        <xsl:when test="(@r:id) and (@w:anchor)">
                            <xsl:attribute name="href">
                                <xsl:value-of select="concat(d:Anchor($myObj,@r:id),'#',@w:anchor)"/>
                            </xsl:attribute>
                            <xsl:attribute name="external">true</xsl:attribute>
                        </xsl:when>
                        <!--If only anchor for hyperlink is present-->
                        <xsl:when test="@w:anchor">
                            <xsl:attribute name="href">
                                <xsl:text>#</xsl:text>
                                <xsl:value-of select="d:EscapeSpecial(@w:anchor)"/>
                            </xsl:attribute>
                        </xsl:when>
                        <!--If only id for hyperlink is present-->
                        <xsl:when test="@r:id">
                            <xsl:attribute name="href">
                                <xsl:value-of select="d:Anchor($myObj,@r:id)"/>
                            </xsl:attribute>
                            <xsl:attribute name="external">true</xsl:attribute>
                        </xsl:when>
                    </xsl:choose>
                    <!--Calling hyperlink template for hyperlink text-->
                    <xsl:call-template name="hyperlink">
                        <xsl:with-param name="verhyp" select="$VERSION"/>
                    </xsl:call-template>
                </a>
            </xsl:if>

            <!--Checking for BookMarkStart-->
            <xsl:if test="name()='w:bookmarkStart'">
                <!--Checking whether BookMarkStart is related to Abbreviations or not -->
                <xsl:if test="substring(@w:name,1,13)='Abbreviations'">
                    <xsl:variable name="full" select="d:FullAbbr($myObj,@w:name,$VERSION)"/>
                    <xsl:choose>
                        <!--Checking whether all previous Abbrevioations tags are closed or not before opening an new Abbreviation tag-->
                        <xsl:when test ="not(d:AbbrAcrFlag($myObj)=1)">
                            <xsl:choose>
                                <!--checking whether an Abbreviation is having Full Form or not-->
                                <xsl:when test="not($full='')">
                                    <xsl:variable name="aquote">"</xsl:variable>
                                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','abbr ','title=',$aquote,$full,$aquote,'&gt;')"/>
                                    <xsl:sequence select="d:sink(d:SetAbbrAcrFlag($myObj))"/> <!-- empty -->
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','abbr','&gt;')"/>
                                    <xsl:sequence select="d:sink(d:SetAbbrAcrFlag($myObj))"/> <!-- empty -->
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:choose>
                                <!--checking whether an Abbreviation is having Full Form or not-->
                                <xsl:when test="not($full='')">
                                    <xsl:variable name="aquote">"</xsl:variable>
                                    <xsl:variable name="temp" select="concat('&lt;','abbr ','title=',$aquote,$full,$aquote,'&gt;')"/>
                                    <xsl:sequence select="d:sink(d:PushAbrAcr($myObj,$temp))"/> <!-- empty -->
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:variable name="aquote">"</xsl:variable>
                                    <xsl:variable name="temp" select="concat('&lt;','abbr','&gt;')"/>
                                    <xsl:sequence select="d:sink(d:PushAbrAcr($myObj,$temp))"/> <!-- empty -->
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:if>
                <!--Checking whether BookMarkStart is related to Acronyms or not -->
                <xsl:if test="substring(@w:name,1,11)='AcronymsYes'">
                    <xsl:variable name="full">
                        <xsl:value-of select="d:FullAcr($myObj,@w:name,$VERSION)"/>
                    </xsl:variable>
                    <xsl:choose>
                        <!--Checking whether all previous Acronyms tags are closed or not before opening an new Acronym tag-->
                        <xsl:when test ="not(d:AbbrAcrFlag($myObj)=1)">
                            <xsl:choose>
                                <!--checking whether an Acronym is having Full Form or not-->
                                <xsl:when test="not($full='')">
                                    <xsl:variable name="aquote">"</xsl:variable>
                                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','acronym ','pronounce=',$aquote,'yes',$aquote,' title=',$aquote,$full,$aquote,'&gt;')"/>
                                    <xsl:sequence select="d:sink(d:SetAbbrAcrFlag($myObj))"/> <!-- empty -->
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:variable name="aquote">"</xsl:variable>
                                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','acronym ','pronounce=',$aquote,'yes',$aquote,'&gt;')"/>
                                    <xsl:sequence select="d:sink(d:SetAbbrAcrFlag($myObj))"/> <!-- empty -->
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:choose>
                                <!--checking whether an Acronym is having Full Form or not-->
                                <xsl:when test="not($full='')">
                                    <xsl:variable name="aquote">"</xsl:variable>
                                    <xsl:variable name="temp" select="concat('&lt;','acronym ','pronounce=',$aquote,'yes',$aquote,' title=',$aquote,$full,$aquote,'&gt;')"/>
                                    <xsl:sequence select="d:sink(d:PushAbrAcr($myObj,$temp))"/> <!-- empty -->
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:variable name="aquote">"</xsl:variable>
                                    <xsl:variable name="temp" select="concat('&lt;','acronym ','pronounce=',$aquote,'yes',$aquote,'&gt;')"/>
                                    <xsl:sequence select="d:sink(d:PushAbrAcr($myObj,$temp))"/> <!-- empty -->
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:if>
                <!--Checking whether BookMarkStart is related to Acronymss or not -->
                <xsl:if test="substring(@w:name,1,10)='AcronymsNo'">
                    <xsl:variable name="full">
                        <xsl:value-of select="d:FullAcr($myObj,@w:name,$VERSION)"/>
                    </xsl:variable>
                    <xsl:choose>
                        <!--Checking whether all previous Acronyms tags are closed or not before opening an new Acronym tag-->
                        <xsl:when test ="not(d:AbbrAcrFlag($myObj)=1)">
                            <xsl:choose>
                                <!--checking whether an Acronym is having Full Form or not-->
                                <xsl:when test="not($full='')">
                                    <xsl:variable name="aquote">"</xsl:variable>
                                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','acronym ','pronounce=',$aquote,'no',$aquote,' title=',$aquote,$full,$aquote,'&gt;')"/>
                                    <xsl:sequence select="d:sink(d:SetAbbrAcrFlag($myObj))"/> <!-- empty -->
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:variable name="aquote">"</xsl:variable>
                                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','acronym ','pronounce=',$aquote,'no',$aquote,'&gt;')"/>
                                    <xsl:sequence select="d:sink(d:SetAbbrAcrFlag($myObj))"/> <!-- empty -->
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:choose>
                                <!--checking whether an Acronym is having Full Form or not-->
                                <xsl:when test="not($full='')">
                                    <xsl:variable name="aquote">"</xsl:variable>
                                    <xsl:variable name="temp" select="concat('&lt;','acronym ','pronounce=',$aquote,'no',$aquote,' title=',$aquote,$full,$aquote,'&gt;')"/>
                                    <xsl:sequence select="d:sink(d:PushAbrAcr($myObj,$temp))"/> <!-- empty -->
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:variable name="aquote">"</xsl:variable>
                                    <xsl:variable name="temp" select="concat('&lt;','acronym ','pronounce=',$aquote,'no',$aquote,'&gt;')"/>
                                    <xsl:sequence select="d:sink(d:PushAbrAcr($myObj,$temp))"/> <!-- empty -->
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:if>

                <!--Checking for hyperlink-->
                <xsl:if test="d:GetHyperlinkName($myObj,@w:name)=1 and not(substring(@w:name,1,13)='Abbreviations') and not(substring(@w:name,1,11)='AcronymsYes') and not(substring(@w:name,1,10)='AcronymsNo')">
                    <xsl:variable name="aquote">"</xsl:variable>
                    <xsl:choose>
                        <!--If hyperling in Table of content-->
                        <xsl:when test="not(contains(@w:name,'_Toc'))">
                            <xsl:sequence select="d:sink(d:TestRun($myObj))"/> <!-- empty -->
                            <xsl:variable name="initialize" select="d:SetHyperLinkFlag($myObj)"/>
                            <xsl:if test="$initialize=1">
                                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','a ','id=',$aquote,d:EscapeSpecial(@w:name),$aquote,'&gt;')"/>
                                <xsl:sequence select="d:sink(d:StroreId($myObj,@w:id))"/> <!-- empty -->
                            </xsl:if>
                            <xsl:if test="$initialize&gt;1">
                                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/a','&gt;')"/>
                                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','a ','id=',$aquote,d:EscapeSpecial(@w:name),$aquote,'&gt;')"/>
                                <xsl:sequence select="d:sink(d:StroreId($myObj,@w:id))"/> <!-- empty -->
                            </xsl:if>
                        </xsl:when>
                    </xsl:choose>
                </xsl:if>

            </xsl:if>

            <!--Checking for BookMarkEnd -->
            <xsl:if test="name()='w:bookmarkEnd'">
                <xsl:variable name="seperate">
                    <xsl:variable name="id" select="@w:id"/>
                    <xsl:choose>
                        <xsl:when test="$flagNote='Note'">
                            <xsl:value-of select="d:BookFootnote($myObj,@w:id)"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:variable name="tempAbbr" select="$documentXml//w:bookmarkStart[@w:id=$id]/@w:name"/>
                            <xsl:value-of select="d:Book($myObj,$tempAbbr)"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:variable>
                <!--Checking whether BookMarkEnd is related to Abbreviations or not -->
                <xsl:if test="$seperate='AbbrTrue'">
                    <!--checking    condition to close abbr Tag -->
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/abbr','&gt;')"/>
                    <xsl:sequence select="d:sink(d:ReSetAbbrAcrFlag($myObj))"/> <!-- empty -->
                    <xsl:if test="d:CountAbrAcr($myObj) &gt; 0">
                        <xsl:value-of disable-output-escaping="yes" select="d:PeekAbrAcr($myObj)"/>
                    </xsl:if>
                </xsl:if>
                <!--Checking whether BookMarkEnd is related to Acronyms or not -->
                <xsl:if test="$seperate='AcrTrue'">
                    <!--checking    condition to close acronym Tag -->
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/acronym','&gt;')"/>
                    <xsl:sequence select="d:sink(d:ReSetAbbrAcrFlag($myObj))"/> <!-- empty -->
                    <xsl:if test="d:CountAbrAcr($myObj) &gt; 0">
                        <xsl:value-of disable-output-escaping="yes" select="d:PeekAbrAcr($myObj)"/>
                    </xsl:if>
                </xsl:if>
                <!--Closing hyperlink if not heading-->
                <xsl:if test="not(d:GetBookmark($myObj)&gt;0)">
                    <xsl:if test="d:CheckId($myObj,@w:id)=1">
                        <xsl:sequence select="d:sink(d:SetTestRun($myObj))"/> <!-- empty -->
                        <xsl:if test="not(../w:pPr/w:pStyle[substring(@w:val,1,7)='Heading'])">
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/a','&gt;')"/>
                        </xsl:if>
                        <xsl:if test="../w:pPr/w:pStyle[substring(@w:val,1,7)='Heading']">
                            <xsl:sequence select="d:sink(d:SetHyperLink($myObj))"/> <!-- empty -->
                        </xsl:if>
                    </xsl:if>
                </xsl:if>
                <xsl:if test="d:GetBookmark($myObj)&gt;0">
                    <xsl:sequence select="d:sink(d:SetTestRun($myObj))"/> <!-- empty -->
                </xsl:if>
            </xsl:if>

            <!--checking sdt element for citation-->
            <xsl:if test="name()='w:sdt'">
                <!--Checking for Citation Element-->
                <xsl:if test="w:sdtContent/w:fldSimple/w:r">
                    <cite>
                        <!--Creating variable SupressAuthor for checking    value '\n'-->
                        <xsl:variable name="SupressAuthor">
                            <xsl:choose>
                                <xsl:when test="./w:sdtContent/w:fldSimple/@w:instr">
                                    <xsl:value-of select="contains(./w:sdtContent/w:fldSimple/@w:instr,'\n')"/>
                                </xsl:when>
                                <xsl:when test="./w:sdtContent/w:r/w:instrText">
                                    <xsl:value-of select="contains(./w:sdtContent/w:r/w:instrText,'\n')"/>
                                </xsl:when>
                            </xsl:choose>
                        </xsl:variable>
                        <!--Creating variable SupressTitle for checking    value '\t'-->
                        <xsl:variable name="SupressTitle">
                            <xsl:choose>
                                <xsl:when test="./w:sdtContent/w:fldSimple/@w:instr">
                                    <xsl:value-of select="contains(./w:sdtContent/w:fldSimple/@w:instr,'\t')"/>
                                </xsl:when>
                                <xsl:when test="./w:sdtContent/w:r/w:instrText">
                                    <xsl:value-of select="contains(./w:sdtContent/w:r/w:instrText,'\t')"/>
                                </xsl:when>
                            </xsl:choose>
                        </xsl:variable>
                        <!--Creating variable SupressYear for checking    value '\y'-->
                        <xsl:variable name="SupressYear">
                            <xsl:choose>
                                <xsl:when test="./w:sdtContent/w:fldSimple/@w:instr">
                                    <xsl:value-of select="contains(./w:sdtContent/w:fldSimple/@w:instr,'\y')"/>
                                </xsl:when>
                                <xsl:when test="./w:sdtContent/w:r/w:instrText">
                                    <xsl:value-of select="contains(./w:sdtContent/w:r/w:instrText,'\y')"/>
                                </xsl:when>
                            </xsl:choose>
                        </xsl:variable>
                        <!--Creating variable TagName to store CitationDetails -->
                        <xsl:variable name="TagName">
                            <xsl:choose>
                                <xsl:when test="./w:sdtContent/w:fldSimple/@w:instr">
                                    <xsl:value-of select="d:CitationDetails($myObj,./w:sdtContent/w:fldSimple/@w:instr)"/>
                                </xsl:when>
                                <xsl:when test="./w:sdtContent/w:r/w:instrText">
                                    <xsl:value-of select="d:CitationDetails($myObj,./w:sdtContent/w:r/w:instrText)"/>
                                </xsl:when>
                            </xsl:choose>
                        </xsl:variable>
                        <xsl:sequence select="d:sink($TagName)"/> <!-- empty -->
                        <xsl:choose>

                            <!--Checking for APA style-->
                            <xsl:when test="$Cite_style='APA' or $Cite_style='GB7714' or $Cite_style='GOST - Name Sort' or $Cite_style='GOST - Title Sort' or $Cite_style='ISO 690 - First Element and Date' or $Cite_style='Turabian' or $Cite_style='Chicago'">
                                <xsl:call-template name="styleCitation">
                                    <xsl:with-param name="supressAuthor" select="$SupressAuthor"/>
                                    <xsl:with-param name="supressTitle" select="$SupressTitle"/>
                                    <xsl:with-param name="supressYear" select="$SupressYear"/>
                                </xsl:call-template>
                            </xsl:when>
                            <!--Checking for MLA style-->
                            <xsl:when test="$Cite_style='MLA'">
                                <xsl:call-template name="styleCitationMLA">
                                    <xsl:with-param name="supressAuthor" select="$SupressAuthor"/>
                                    <xsl:with-param name="supressTitle" select="$SupressTitle"/>
                                    <xsl:with-param name="supressYear" select="$SupressYear"/>
                                </xsl:call-template>
                            </xsl:when>
                            <!--Checking forISO 690 - Numerical Reference style-->
                            <xsl:when test="$Cite_style='ISO 690 - Numerical Reference'">
                                <xsl:value-of select="./w:sdtContent//w:t"/>
                            </xsl:when>
                            <!--Checking for SIST02 style-->
                            <xsl:when test="$Cite_style='SIST02'">
                                <xsl:call-template name="styleCitationSIST02">
                                    <xsl:with-param name="supressAuthor" select="$SupressAuthor"/>
                                    <xsl:with-param name="supressTitle" select="$SupressTitle"/>
                                    <xsl:with-param name="supressYear" select="$SupressYear"/>
                                </xsl:call-template>
                            </xsl:when>
                        </xsl:choose>
                    </cite>
                </xsl:if>
            </xsl:if>

            <xsl:if test="name()='w:del'">
                <xsl:if test="$prmTrack='No'">
                    <xsl:for-each select="w:r">
                        <xsl:message terminate="no">progress:parahandler</xsl:message>
                        <xsl:value-of select="w:delText"/>
                    </xsl:for-each>
                </xsl:if>
            </xsl:if>

            <xsl:if test="name()='w:ins'">
                <xsl:if test="$prmTrack='Yes'">
                    <xsl:for-each select="w:r">
                        <xsl:message terminate="no">progress:parahandler</xsl:message>
                        <xsl:value-of select="w:t"/>
                    </xsl:for-each>
                </xsl:if>
            </xsl:if>

            <xsl:if test="name()='w:r' and not(preceding-sibling::w:pPr/w:pStyle[substring(@w:val,1,3)='TOC'])">
                <xsl:call-template name="TempParaHandler">
                    <xsl:with-param name="flag" select="$flag"/>
                    <xsl:with-param name="imgOptionPara" select="$imgOptionPara"/>
                    <xsl:with-param name="dpiPara" select="$dpiPara"/>
                    <xsl:with-param name="custom" select="$custom"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="charparahandlerStyle" select="$charparahandlerStyle"/>
                    <xsl:with-param name="VERSION" select="$VERSION"/>
                </xsl:call-template>
            </xsl:if>

            <!--Capturing Fidelity loss for Copy Right-->
            <xsl:if test="name()='w:sdt'">
                <xsl:if test="not(w:sdtPr/w:citation)">
                    <xsl:message terminate="no">
                        <xsl:value-of select="concat('translation.oox2Daisy.UncoveredElement|','Copy Right')"/>
                    </xsl:message>
                </xsl:if>
            </xsl:if>

        </xsl:for-each>

        <!--checking    condition to close bdo Tag -->
        <xsl:if test="d:SetbdoFlag($myObj)&gt;=2">
            <!--Closing BDO tag-->
            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
        </xsl:if>

        <!--reseting bdo flag-->
        <xsl:sequence select="d:sink(d:reSetbdoFlag($myObj))"/> <!-- empty -->
        <xsl:sequence select="d:sink(d:AssingBookmark($myObj))"/> <!-- empty -->

        <!--closing paragraph tag-->
        <xsl:if test="not($flag='0') and not(d:AbbrAcrFlag($myObj)=1) and not($flagNote='hyper')">
            <xsl:if test="not(d:GetTestRun($myObj)&gt;='1') and (d:GetCodeFlag($myObj)='0') and (not(d:Getlinenumflag($myObj)=0))">
                <xsl:choose>
                    <!--check if <p> teg was opened in "Languages" template-->
                    <xsl:when test="(w:r/w:rPr/w:lang) or (w:r/w:rPr/w:rFonts/@w:hint)">
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:if test="(not(w:r/w:rPr/w:rStyle[@w:val='PageNumberDAISY']) and ($custom='Custom')) or (not($custom='Custom'))">
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
                        </xsl:if>
                    </xsl:otherwise>
                </xsl:choose>
                <xsl:if test="(d:ListMasterSubFlag($myObj)=1) and ($mastersubpara = 'Yes')">
                    <xsl:variable name="curLevel" select="d:PeekLevel($myObj)"/>
                    <xsl:value-of disable-output-escaping="yes" select="d:ClosingMasterSub($myObj,$curLevel)"/>
                    <xsl:value-of disable-output-escaping="yes" select="d:PeekMasterSubdoc($myObj)"/>
                    <xsl:sequence select="d:sink(d:MasterSubResetFlag($myObj))"/> <!-- empty -->
                    <xsl:value-of disable-output-escaping="yes" select="d:OpenMasterSub($myObj,$curLevel)"/>
                </xsl:if>
            </xsl:if>
        </xsl:if>

        <!--Checking for heading flag and abbr flag and closing paragraph tag-->
        <xsl:if test="not($flag='0') and d:AbbrAcrFlag($myObj)=1">
            <xsl:variable name="temp" select="concat('&lt;','/p','&gt;')"/>
            <xsl:sequence select="d:sink(d:PushAbrAcrpara($myObj,$temp))"/> <!-- empty -->
        </xsl:if>
        <xsl:sequence select="d:sink(d:SetGetHyperLinkFlag($myObj))"/> <!-- empty -->
        <xsl:sequence select="d:sink(d:ReSetListFlag($myObj))"/> <!-- empty -->
        <!--code for hard return-->
        <xsl:text>&#10;</xsl:text>

    </xsl:template>

    <xsl:template name="TempSidebar">
        <xsl:param name="flag"/>
        <xsl:param name="custom"/>
        <xsl:param name="txt"/>
        <xsl:param name="prmTrack"/>
        <xsl:param name="VERSION"/>
        <xsl:param name="charparahandlerStyle"/>
        <xsl:param name="mastersubpara"/>
        <xsl:param name="level"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>

        <xsl:if test="$flag=0">
            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;',concat('/h',$level),'&gt;')"/>
        </xsl:if>
        <xsl:for-each select="w:r/w:pict//v:textbox/w:txbxContent">
            <xsl:message terminate="no">progress:parahandler</xsl:message>
            <sidebar>
                <xsl:attribute    name="render">required</xsl:attribute>
                <xsl:for-each select="./node()">
                    <xsl:message terminate="no">progress:parahandler</xsl:message>
                    <xsl:choose>
                        <!--Checking for Headings in sidebar-->
                        <xsl:when test="(w:pPr/w:pStyle[substring(@w:val,1,7)='Heading']) or (w:pPr/w:pStyle/@w:val='BridgeheadDAISY')">
                            <hd>
                                <xsl:call-template name="ParaHandler">
                                    <xsl:with-param name="flag" select="'0'"/>
                                    <xsl:with-param name="txt" select="$txt"/>
                                    <xsl:with-param name="custom" select="$custom"/>
                                    <xsl:with-param name="charparahandlerStyle" select="$charparahandlerStyle"/>
                                </xsl:call-template>
                            </hd>
                        </xsl:when>
                        <!--Checking for lists in sidebar-->
                        <xsl:when test="((w:pPr/w:numPr/w:ilvl) and (w:pPr/w:numPr/w:numId))">
                            <xsl:call-template name="List">
                                <xsl:with-param name="prmTrack" select="$prmTrack"/>
                                <xsl:with-param name="verlist" select="$VERSION"/>
                                <xsl:with-param name="listcharStyle" select="$charparahandlerStyle"/>
                            </xsl:call-template>
                        </xsl:when>
                        <!--Checking for Table in sidebar-->
                        <xsl:when test="name()='w:tbl'">
                            <xsl:call-template name="TableHandler">
                                <xsl:with-param name="parmVerTable" select="$VERSION"/>
                                <xsl:with-param name="custom" select="$custom"/>
                                <xsl:with-param name="mastersubtbl" select="$mastersubpara"/>
                                <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                            </xsl:call-template>
                        </xsl:when>
                        <!--Checking for Prodnote style in sidebar-->
                        <xsl:when test="(w:pPr/w:pStyle/@w:val='Prodnote-RequiredDAISY') or (w:pPr/w:pStyle/@w:val='Prodnote-OptionalDAISY')">
                            <xsl:call-template name="ParagraphStyle">
                                <xsl:with-param name="prmTrack" select="$prmTrack"/>
                                <xsl:with-param name="VERSION" select="$VERSION"/>
                                <xsl:with-param name="custom" select="$custom"/>
                                <xsl:with-param name="flag" select="'0'"/>
                                <xsl:with-param name="txt" select="$txt"/>
                                <xsl:with-param name="characterparaStyle" select="$charparahandlerStyle"/>
                            </xsl:call-template>
                        </xsl:when>


                        <xsl:otherwise>
                            <xsl:if test="not(w:pPr/w:pStyle/@w:val='List-HeadingDAISY')">
                                <!--Calling StyleContainer Template -->
                                <xsl:call-template name="StyleContainer">
                                    <xsl:with-param name="VERSION" select="$VERSION"/>
                                    <xsl:with-param name="prmTrack" select="$prmTrack"/>
                                    <xsl:with-param name="custom" select="$custom"/>
                                    <xsl:with-param name="mastersubstyle" select="$mastersubpara"/>
                                    <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                                </xsl:call-template>
                            </xsl:if>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:for-each>
            </sidebar>
        </xsl:for-each>
    </xsl:template>

    <xsl:template name="TempParaHandler">
        <xsl:param name="flag"/>
        <xsl:param name="imgOptionPara"/>
        <xsl:param name="dpiPara"/>
        <xsl:param name="custom"/>
        <xsl:param name="txt"/>
        <xsl:param name="charparahandlerStyle"/>
        <xsl:param name="VERSION"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:variable name="quote">"</xsl:variable>

        <!--Checking for line breaks-->
        <xsl:if test="((w:br/@w:type='textWrapping') or (w:br)) and (not(w:br/@w:type='page'))">
            <br/>
        </xsl:if>

        <xsl:if test="w:tab">
            <xsl:text> </xsl:text>
        </xsl:if>

        <!--Checking for page breaks and populating page numbers.-->
        <!-- DB : if list skip this-->
        <xsl:if test="((w:lastRenderedPageBreak) or (w:br/@w:type='page'))
                                                                                        and not($flag='0')
                                                                                        and ($custom='Automatic')
                                                                                        and not(((../w:pPr/w:numPr/w:ilvl)
                                                                                                                and (../w:pPr/w:numPr/w:numId)
                                                                                                                and not(../w:pPr/w:rPr/w:vanish))
                                                                                                        and not(../w:pPr/w:pStyle[substring(@w:val,1,7)='Heading']))">
            <xsl:if test="not($flag='2')">
                <xsl:choose>
                    <xsl:when test="not(w:t) and (w:lastRenderedPageBreak) and (w:br/@w:type='page')">
                        <xsl:if test="not(../following-sibling::w:sdt[1]/w:sdtPr/w:docPartObj/w:docPartGallery/@w:val='Table of Contents')">
                            <xsl:if test="not(../preceding-sibling::node()[1]/w:pPr/w:sectPr)">
                                <xsl:sequence select="d:sink(d:IncrementPage($myObj))"/> <!-- empty -->
                                <!--Closing paragraph tag-->
                                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
                                <xsl:if test="$flag='3'">
                                    <!--Closing paragraph tag-->
                                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p','&gt;')"/>
                                </xsl:if>
                                <!--calling template to initialize page number information-->
                                <xsl:call-template name="SectionBreak">
                                    <xsl:with-param name="count" select="'1'"/>
                                    <xsl:with-param name="node" select="'body'"/>
                                </xsl:call-template>
                                <!--producer note for blank pages-->
                                <prodnote>
                                    <xsl:attribute name="render">optional</xsl:attribute>
                                    <xsl:value-of select="'Blank Page'"/>
                                </prodnote>
                                <xsl:if test="$flag='3'">
                                    <!--Closing paragraph tag-->
                                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
                                </xsl:if>
                                <!--Opening paragraph tag-->
                                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p','&gt;')"/>
                            </xsl:if>
                        </xsl:if>
                    </xsl:when>
                    <!--Checking for page breaks and populating page numbers.-->
                    <xsl:when test="((w:br/@w:type='page')
                                                                                                                        and not((../following-sibling::w:p[1]/w:pPr/w:sectPr)
                                                                                                                                        or (following-sibling::w:r[1]/w:lastRenderedPageBreak)
                                                                                                                                        or (../following-sibling::w:p[2]/w:r/w:lastRenderedPageBreak)
                                                                                                                                        or (../following-sibling::w:p[1]/w:r/w:lastRenderedPageBreak)))">
                        <!--Incrementing page numbers-->
                        <xsl:sequence select="d:sink(d:IncrementPage($myObj))"/> <!-- empty -->
                        <!--Closing paragraph tag-->
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
                        <xsl:if test="$flag='3'">
                            <!--Opening paragraph tag-->
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p','&gt;')"/>
                        </xsl:if>
                        <!--calling template to initialize page number information-->
                        <xsl:call-template name="SectionBreak">
                            <xsl:with-param name="count" select="'1'"/>
                            <xsl:with-param name="node" select="'body'"/>
                        </xsl:call-template>
                        <xsl:if test="$flag='3'">
                            <!--Closing paragraph tag-->
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
                        </xsl:if>
                        <!--Opening paragraph tag-->
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p','&gt;')"/>
                    </xsl:when>
                    <xsl:when test="(w:lastRenderedPageBreak)
                                                                                                                and not(../w:pPr/w:sectPr                                        
                                                                                                                                or ../w:pPr/w:pStyle[substring(@w:val,1,5)='Index'])">
                        <xsl:sequence select="d:sink(d:IncrementPage($myObj))"/> <!-- empty -->
                        <!--Closing paragraph tag-->
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
                        <xsl:if test="$flag='3'">
                            <!--Opening paragraph tag-->
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p','&gt;')"/>
                        </xsl:if>
                        <!--calling template to initialize page number information-->
                        <xsl:call-template name="SectionBreak">
                            <xsl:with-param name="count" select="'1'"/>
                            <xsl:with-param name="node" select="'body'"/>
                        </xsl:call-template>
                        <xsl:if test="$flag='3'">
                            <!--Closing paragraph tag-->
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
                        </xsl:if>
                        <!--Opening paragraph tag-->
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p','&gt;')"/>
                    </xsl:when>
                </xsl:choose>
            </xsl:if>
        </xsl:if>

        <xsl:choose>
            <!--Checking for Images-->
            <xsl:when test="w:pict">
                <xsl:choose>
                    <!--checking for Images in word2003 version-->
                    <xsl:when test="w:pict/v:shape/v:imagedata/@r:id">
                        <xsl:call-template name="Imagegroup2003">
                            <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                        </xsl:call-template>
                    </xsl:when>
                    <!--checking image groups-->
                    <xsl:when test="(w:pict/v:group) and (($VERSION='11.0') or ($VERSION='10.0'))and (not(descendant::node()[name()='w:txbxContent'])) and (not(w:pict/v:rect/v:textbox/w:txbxContent))">
                        <xsl:call-template name="Imagegroups">
                            <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                        </xsl:call-template>
                    </xsl:when>
                    <xsl:when test="(w:pict/v:group) and ($VERSION='12.0')">
                        <xsl:call-template name="Imagegroups">
                            <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                        </xsl:call-template>
                    </xsl:when>
                    <xsl:when test="(w:pict/v:shape/@o:spid)    and (not(descendant::node()[name()='w:txbxContent'])) and(not(w:pict/v:rect/v:textbox/w:txbxContent))">
                        <xsl:call-template name="tmpShape">
                            <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                        </xsl:call-template>
                    </xsl:when>
                    <xsl:when test="(w:pict/v:shape) and not(contains(w:pict/v:shape/@id,'i')) and (not(descendant::node()[name()='w:txbxContent'])) and(not(w:pict/v:rect/v:textbox/w:txbxContent))">
                        <xsl:call-template name="tmpShape">
                            <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                        </xsl:call-template>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:choose>
                            <xsl:when test="not(w:pict/v:shape/v:textbox) and not(contains(w:pict/v:shape/@id,'i')) and(not(w:pict/v:rect/v:textbox/w:txbxContent)) and (not(($VERSION='11.0') or ($VERSION='10.0'))) ">
                                <xsl:call-template name="tmpShape">
                                    <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                                </xsl:call-template>
                            </xsl:when>
                            <xsl:when test="not(w:pict/v:shape/v:textbox) and not(contains(w:pict/v:shape/@id,'i')) and(not(w:pict/v:rect/v:textbox/w:txbxContent)) and($VERSION='12.0')">
                                <xsl:call-template name="tmpShape">
                                    <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                                </xsl:call-template>
                            </xsl:when>
                        </xsl:choose>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>

            <!--checking for images in word2007 version-->
            <xsl:when test="w:drawing">
                <xsl:call-template name="PictureHandler">
                    <xsl:with-param name="imgOpt" select="$imgOptionPara"/>
                    <xsl:with-param name="dpi" select="$dpiPara"/>
                    <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                </xsl:call-template>
            </xsl:when>

            <!--Checking for objects-->
            <xsl:when test="w:object/o:OLEObject">
                <xsl:choose>
                    <!--Checking for Design science Math Equations-->
                    <xsl:when test="w:object/o:OLEObject[@ProgID='Equation.DSMT4']">

                        <xsl:variable name="Math_DSMT4">
                            <xsl:choose>
                                <xsl:when test="ancestor::node()[name()='w:txbxContent']">
                                    <xsl:value-of select="d:GetMathML($myObj,'wdTextFrameStory')"/>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:value-of select="d:GetMathML($myObj,'wdMainTextStory')"/>
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:variable>

                        <xsl:choose>
                            <xsl:when test="$Math_DSMT4=''">
                                <imggroup>
                                    <img>
                                        <!--Creating variable mathimage for storing r:id value from document.xml-->
                                        <xsl:variable name="Math_rid">
                                            <xsl:value-of select="w:object/v:shape/v:imagedata/@r:id"/>
                                        </xsl:variable>
                                        <xsl:attribute name="alt">
                                            <xsl:choose>
                                                <!--Checking for alt Text-->
                                                <xsl:when test="w:object/v:shape/@alt">
                                                    <xsl:value-of select="w:object/v:shape/@alt"/>
                                                </xsl:when>
                                                <xsl:otherwise>
                                                    <!--Hardcoding value 'Math Equation'if user donot provide alt text for Math Equations-->
                                                    <xsl:value-of select="'Math Equation'"/>
                                                </xsl:otherwise>
                                            </xsl:choose>
                                        </xsl:attribute>
                                        <xsl:attribute name="src">
                                            <!--Calling MathImage function-->
                                            <xsl:value-of select="d:MathImage($myObj,$Math_rid)"/>
                                        </xsl:attribute>
                                    </img>
                                </imggroup>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of disable-output-escaping="yes" select="$Math_DSMT4"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:when>
                    <!--Checking condition for MathEquations in word 2003/xp-->
                    <xsl:when test="contains(w:object/o:OLEObject/@ProgID,'Equation')and not(w:object/o:OLEObject[@ProgID='Equation.DSMT4'])">
                        <xsl:variable name="mathimage">
                            <xsl:value-of select="w:object/v:shape/v:imagedata/@r:id"/>
                        </xsl:variable>
                        <imggroup>
                            <img>
                                <!--<xsl:value-of select="$mathimage"/>-->
                                <xsl:attribute name="alt">
                                    <xsl:choose>
                                        <xsl:when test="w:object/v:shape/@alt">
                                            <xsl:value-of select="w:object/v:shape/@alt"/>
                                        </xsl:when>
                                        <xsl:otherwise>
                                            <xsl:value-of select="'Math Equation'"/>
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </xsl:attribute>
                                <xsl:attribute name="src">
                                    <xsl:value-of select="d:MathImage($myObj,$mathimage)"/>
                                </xsl:attribute>
                            </img>
                        </imggroup>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:call-template name="Object">
                            <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                        </xsl:call-template>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>

            <!--Checking for footnotes in word2003 and word2007-->
            <xsl:when test="(w:rPr/w:rStyle[@w:val ='FootnoteReference']) or (w:footnoteReference)">
                <xsl:choose>
                    <xsl:when test="w:rPr/w:rStyle[@w:val ='FootnoteReference']">
                        <xsl:variable name="class" select="w:rPr/w:rStyle/@w:val"/>
                        <xsl:variable name="footnoteid" select="w:footnoteReference/@w:id"/>
                        <xsl:sequence select="d:sink(d:AddFootNote($myObj,$footnoteid))"/> <!-- empty -->
                        <xsl:call-template name="tmpProcessFootNote">
                            <xsl:with-param name="varFootnote_Id" select="$footnoteid"/>
                            <xsl:with-param name="varNote_Class" select="$class"/>
                            <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                        </xsl:call-template>
                    </xsl:when>
                    <!---->
                    <xsl:when test="w:footnoteReference">
                        <xsl:variable name="class" select="'FootnoteReference'"/>
                        <xsl:variable name="footnoteid" select="w:footnoteReference/@w:id"/>
                        <xsl:sequence select="d:sink(d:AddFootNote($myObj,$footnoteid))"/> <!-- empty -->
                        <xsl:call-template name="tmpProcessFootNote">
                            <xsl:with-param name="varFootnote_Id" select="$footnoteid"/>
                            <xsl:with-param name="varNote_Class" select="$class"/>
                            <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                        </xsl:call-template>
                    </xsl:when>
                </xsl:choose>
            </xsl:when>

            <!--Checking for endnotes in word2003 and word2007-->
            <xsl:when test="(w:rPr/w:rStyle[@w:val='EndnoteReference']) or (w:endnoteReference)">
                <xsl:choose>
                    <xsl:when test="w:rPr/w:rStyle[@w:val='EndnoteReference']">
                        <xsl:variable name="class" select="w:rPr/w:rStyle/@w:val"/>
                        <xsl:variable name="endnoteid" select="w:endnoteReference/@w:id"/>
                        <xsl:call-template name="tmpProcessFootNote">
                            <xsl:with-param name="varFootnote_Id" select="$endnoteid"/>
                            <xsl:with-param name="varNote_Class" select="$class"/>
                            <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                        </xsl:call-template>
                    </xsl:when>
                    <xsl:when test="w:endnoteReference">
                        <xsl:variable name="class" select="'EndnoteReference'"/>
                        <xsl:variable name="endnoteid" select="w:endnoteReference/@w:id"/>
                        <xsl:call-template name="tmpProcessFootNote">
                            <xsl:with-param name="varFootnote_Id" select="$endnoteid"/>
                            <xsl:with-param name="varNote_Class" select="$class"/>
                            <xsl:with-param name="characterStyle" select="$charparahandlerStyle"/>
                        </xsl:call-template>
                    </xsl:when>
                </xsl:choose>
            </xsl:when>

            <xsl:otherwise>
                <xsl:choose>
                    <!--Checking for BDO Element-->
                    <xsl:when test="(../w:pPr/w:bidi) and (w:rPr/w:rtl)">
                        <xsl:variable name="Bd">
                            <xsl:call-template name="BdoRtlLanguages"/>
                        </xsl:variable>
                        <xsl:variable name="bdoflag" select="d:SetbdoFlag($myObj)"/>
                        <!--opening bdo Tag-->
                        <xsl:if test="$bdoflag=1">
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
                        </xsl:if>
                    </xsl:when>

                    <!--Checking for BDO Element-->
                    <xsl:when test="(../w:pPr/w:bidi)">
                        <xsl:variable name="Bd">
                            <!--Calling template for Bdo language-->
                            <xsl:call-template name="BdoRtlLanguages"/>
                        </xsl:variable>
                        <!--Setting flag for bdo-->
                        <xsl:variable name="bdoflag" select="d:SetbdoFlag($myObj)"/>
                        <!--checking condition for opening p tag-->
                        <!--opening bdo Tag-->
                        <xsl:if test="$bdoflag=1">
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
                        </xsl:if>
                    </xsl:when>

                    <!--Checking for BDO Element-->
                    <xsl:when test="w:rPr/w:rtl">
                        <xsl:variable name="Bd">
                            <!--Calling template for Bdo language-->
                            <xsl:call-template name="BdoRtlLanguages"/>
                        </xsl:variable>
                        <!--Setting flag for bdo-->
                        <xsl:variable name="bdoflag" select="d:SetbdoFlag($myObj)"/>
                        <!--opening bdo Tag-->
                        <xsl:if test="$bdoflag=1">
                            <!--If flag value is 1 opening bdo element-->
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
                        </xsl:if>
                    </xsl:when>
                </xsl:choose>

                <!--Calling Custom styles template if not caption-->
                <xsl:if test="not((((../w:pPr/w:pStyle/@w:val='Table-CaptionDAISY')
                                                                                                                                or (../w:pPr/w:pStyle/@w:val='Caption'))
                                                                                                                and ((../following-sibling::node()[1][name()='w:tbl'])
                                                                                                                                or (../preceding-sibling::node()[1][name()='w:tbl'])))
                                                                                                 or ((../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')
                                                                                                                and ((../following-sibling::node()[1]/w:r/w:drawing)
                                                                                                                                or (../following-sibling::node()[1]/w:r/w:pict)
                                                                                                                                or (../following-sibling::node()[1]/w:r/w:object)
                                                                                                                                or (../w:r/w:drawing)
                                                                                                                                or (../w:r/w:pict)
                                                                                                                                or (../w:r/w:object))))">
                    <xsl:choose>
                        <xsl:when test="d:ListFlag($myObj)=0">
                            <xsl:sequence select="d:sink(d:SetListFlag($myObj))"/> <!-- empty -->
                            <xsl:call-template name="CustomStyles">
                                <xsl:with-param name="customTag" select="w:rPr/w:rStyle/@w:val"/>
                                <xsl:with-param name="custom" select="$custom"/>
                                <xsl:with-param name="txt" select="$txt"/>
                                <xsl:with-param name="customcharStyle" select="$charparahandlerStyle"/>
                            </xsl:call-template>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:call-template name="CustomStyles">
                                <xsl:with-param name="customTag" select="w:rPr/w:rStyle/@w:val"/>
                                <xsl:with-param name="custom" select="$custom"/>
                                <xsl:with-param name="customcharStyle" select="$charparahandlerStyle"/>
                            </xsl:call-template>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:if>
            </xsl:otherwise>
        </xsl:choose>

        <!--Initializing Hyperlink flag-->
        <xsl:sequence select="d:sink(d:SetGetHyperLinkFlag($myObj))"/> <!-- empty -->
        <!--Closing hyperlink tag for headings-->
        <xsl:if test="(d:GetFlag($myObj)&gt;=1) and (../w:pPr/w:pStyle[substring(@w:val,1,7)='Heading'])">
            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/a','&gt;')"/>
            <xsl:sequence select="d:sink(d:GetHyperLink($myObj))"/> <!-- empty -->
        </xsl:if>

    </xsl:template>

    <!--Template for hyperlink-->
    <xsl:template name="hyperlink">
        <xsl:param name="verhyp"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:for-each select="w:r">
            <xsl:message terminate="no">progress:parahandler</xsl:message>
            <xsl:call-template name="CustomCharStyle"/>
        </xsl:for-each>
    </xsl:template>

    <!--Template for smartTag-->
    <xsl:template name="smartTag">
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:for-each select="w:r">
            <xsl:message terminate="no">progress:parahandler</xsl:message>
            <xsl:value-of select="w:t"/>
        </xsl:for-each>
        <xsl:for-each select="w:smartTag">
            <xsl:message terminate="no">progress:parahandler</xsl:message>
            <xsl:for-each select="w:r">
                <xsl:message terminate="no">progress:parahandler</xsl:message>
                <xsl:value-of select="w:t"/>
            </xsl:for-each>
        </xsl:for-each>
    </xsl:template>

    <!--Template for fldSimple-->
    <xsl:template name="fldSimple">
        <xsl:param name="characterStyle"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:for-each select="w:r">
            <xsl:message terminate="no">progress:parahandler</xsl:message>
            <xsl:value-of select="w:t"/>
        </xsl:for-each>
    </xsl:template>

    <!--template for traping fidelity loss element structure document-->
    <xsl:template name="structuredocument">
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:if test="w:sdtPr/w:docPartObj/w:docPartGallery">
            <!--Displaying fidelity loss message-->
            <xsl:message terminate="no">
                <xsl:value-of select="concat('translation.oox2Daisy.sdtElement|',w:sdtPr/w:docPartObj/w:docPartGallery/@w:val)"/>
            </xsl:message>
        </xsl:if>
    </xsl:template>

    <!--Template to implement citation styles-->
    <xsl:template name="styleCitation">
        <xsl:param name="supressAuthor"/>
        <xsl:param name="supressTitle"/>
        <xsl:param name="supressYear"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:choose>
            <!--Checking condition for supressing Author,Title,Year-->
            <xsl:when test="($supressAuthor='true')and ($supressTitle='true') and($supressYear='true')">
                <xsl:value-of select="./w:sdtContent//w:t"/>
            </xsl:when>
            <!--Checking condition for supressing Author,Title-->
            <xsl:when test="($supressAuthor='true') and ($supressTitle='true')">
                <xsl:value-of select="d:GetYear($myObj)"/>
            </xsl:when>
            <!--Checking condition for supressing Author,Year-->
            <xsl:when test="($supressAuthor='true') and ($supressYear='true')">
                <xsl:text>(</xsl:text>
                <title>
                    <xsl:value-of select="d:GetTitle($myObj)"/>
                </title>
                <xsl:text>)</xsl:text>
            </xsl:when>
            <!--Checking condition for supressing Title,Year-->
            <xsl:when test="($supressTitle='true') and ($supressYear='true')">
                <xsl:text>(</xsl:text>
                <author>
                    <xsl:value-of select="d:GetAuthor($myObj)"/>
                </author>
                <xsl:text>)</xsl:text>
            </xsl:when>
            <!--Checking condition for supressing Author-->
            <xsl:when test="$supressAuthor='true'">
                <xsl:text>(</xsl:text>
                <title>
                    <xsl:value-of select="d:GetTitle($myObj)"/>
                </title>
                <xsl:text>,</xsl:text>
                <xsl:value-of select="d:GetYear($myObj)"/>
                <xsl:text>)</xsl:text>
            </xsl:when>
            <!--Checking condition for supressing Title-->
            <xsl:when test="$supressTitle='true'">
                <xsl:text>(</xsl:text>
                <author>
                    <xsl:value-of select="d:GetAuthor($myObj)"/>
                </author>
                <xsl:text>,</xsl:text>
                <xsl:value-of select="d:GetYear($myObj)"/>
                <xsl:text>)</xsl:text>
            </xsl:when>
            <!--Checking condition for supressing Year-->
            <xsl:when test="$supressYear='true'">
                <xsl:text>(</xsl:text>
                <author>
                    <!--Calling GetAuthor function to get the value of Author-->
                    <xsl:value-of select="d:GetAuthor($myObj)"/>
                </author>
                <xsl:text>)</xsl:text>
            </xsl:when>
            <xsl:otherwise>
                <xsl:choose>
                    <xsl:when test="d:GetAuthor($myObj)=''">
                        <xsl:value-of select="./w:sdtContent//w:t"/>
                    </xsl:when>
                    <xsl:when test="(d:GetAuthor($myObj)='') and (d:GetYear($myObj)='') ">
                        <title>
                            <!--Calling GetTitle function to get the value of Title-->
                            <xsl:value-of select="d:GetTitle($myObj)"/>
                        </title>
                    </xsl:when>
                    <xsl:when test="(d:GetAuthor($myObj)='') and (d:GetTitle($myObj)='') and (d:GetYear($myObj)='')">
                        <xsl:value-of select="./w:sdtContent//w:t"/>
                    </xsl:when>
                    <xsl:when test="(d:GetTitle($myObj)='') and (d:GetYear($myObj)='')">
                        <author>
                            <xsl:text>(</xsl:text>
                            <!--Calling GetAuthor function to get the value of Author-->
                            <xsl:value-of select="d:GetAuthor($myObj)"/>
                            <xsl:text>)</xsl:text>
                        </author>
                    </xsl:when>
                    <xsl:when test="(d:GetAuthor($myObj)='') and (d:GetTitle($myObj)='')">
                        <!--Calling GetYear function to get the value of Year-->
                        <xsl:value-of select="d:GetYear($myObj)"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:text>(</xsl:text>
                        <author>
                            <!--Calling GetAuthor function to get the value of Author-->
                            <xsl:value-of select="d:GetAuthor($myObj)"/>
                        </author>
                        <xsl:text>,</xsl:text>
                        <!--Calling GetYear function to get the value of the Year-->
                        <xsl:value-of select="d:GetYear($myObj)"/>
                        <xsl:text>)</xsl:text>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <!--Template to implement citation styles MLA-->
    <xsl:template name="styleCitationMLA">
        <xsl:param name="supressAuthor"/>
        <xsl:param name="supressTitle"/>
        <xsl:param name="supressYear"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:choose>
            <!--Checking condition for supressing Author,Title,Year-->
            <xsl:when test="($supressAuthor='true')and ($supressTitle='true') and($supressYear='true')">
                <xsl:value-of select="./w:sdtContent//w:t"/>
            </xsl:when>
            <xsl:when test="$supressAuthor='true'">
                <title>
                    <!--Calling GetTitle function to get the value of the Title-->
                    <xsl:value-of select="d:GetTitle($myObj)"/>
                </title>
            </xsl:when>
            <xsl:when test="$supressTitle='true'">
                <author>
                    <!--Calling GetAuthor function to get the value of the Author-->
                    <xsl:value-of select="d:GetAuthor($myObj)"/>
                </author>
            </xsl:when>
            <xsl:otherwise>
                <xsl:choose>
                    <xsl:when test="d:GetAuthor($myObj)=''">
                        <xsl:value-of select="./w:sdtContent//w:t"/>
                    </xsl:when>
                    <xsl:when test="(d:GetAuthor($myObj)='') and (d:GetTitle($myObj)='') and (d:GetYear($myObj)='')">
                        <xsl:value-of select="./w:sdtContent//w:t"/>
                    </xsl:when>
                    <xsl:when test="(d:GetTitle($myObj)='')">
                        <author>
                            <xsl:text>(</xsl:text>
                            <xsl:value-of select="d:GetAuthor($myObj)"/>
                            <xsl:text>)</xsl:text>
                        </author>
                    </xsl:when>
                    <xsl:otherwise>
                        <author>
                            <xsl:text>(</xsl:text>
                            <xsl:value-of select="d:GetAuthor($myObj)"/>
                            <xsl:text>)</xsl:text>
                        </author>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <!--Template to implement citation styleSIST02-->
    <xsl:template name="styleCitationSIST02">
        <xsl:param name="supressAuthor"/>
        <xsl:param name="supressTitle"/>
        <xsl:param name="supressYear"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:choose>
            <!--Checking condition for supressing Author,Title,Year-->
            <xsl:when test="($supressAuthor='true')and ($supressTitle='true') and($supressYear='true')">
                <xsl:for-each select="./w:sdtContent/w:fldSimple/w:r">
                    <xsl:message terminate="no">progress:parahandler</xsl:message>
                    <xsl:value-of select="w:t"/>
                </xsl:for-each>
            </xsl:when>
            <xsl:when test="($supressAuthor='true')">
                <xsl:text>(</xsl:text>
                <!--Calling GetYear function to get the value of the Year-->
                <xsl:value-of select="d:GetYear($myObj)"/>
                <xsl:text>)</xsl:text>
            </xsl:when>
            <!--Checking condition for supressing Author,Year-->
            <xsl:when test="($supressAuthor='true') and ($supressYear='true')">
                <xsl:value-of select="./w:sdtContent//w:t"/>
            </xsl:when>
            <!--Checking condition for supressing Title,Year-->
            <xsl:when test="($supressTitle='true') and ($supressYear='true')">
                <author>
                    <xsl:text>(</xsl:text>
                    <!--Calling GetAuthor function to get the value of the Author-->
                    <xsl:value-of select="d:GetAuthor($myObj)"/>
                    <xsl:text>)</xsl:text>
                </author>
            </xsl:when>
            <!--Checking condition for supressing Author-->
            <xsl:when test="$supressAuthor='true'">
                <!--Calling GetYear function to get the value of the Year-->
                <xsl:value-of select="d:GetYear($myObj)"/>
            </xsl:when>
            <!--Checking condition for supressing Year-->
            <xsl:when test="$supressYear='true'">
                <author>
                    <xsl:text>(</xsl:text>
                    <!--Calling GetAuthor function to get the value of the Author-->
                    <xsl:value-of select="d:GetAuthor($myObj)"/>
                    <xsl:text>)</xsl:text>
                </author>
            </xsl:when>
            <xsl:otherwise>
                <xsl:choose>
                    <xsl:when test="d:GetAuthor($myObj)=''">
                        <xsl:for-each select="./w:sdtContent/w:fldSimple/w:r">
                            <xsl:message terminate="no">progress:parahandler</xsl:message>
                            <xsl:value-of select="w:t"/>
                        </xsl:for-each>
                    </xsl:when>
                    <xsl:when test="(d:GetAuthor($myObj)='') and (d:GetTitle($myObj)='') and (d:GetYear($myObj)='')">
                        <xsl:for-each select="./w:sdtContent/w:fldSimple/w:r">
                            <xsl:message terminate="no">progress:parahandler</xsl:message>
                            <xsl:value-of select="w:t"/>
                        </xsl:for-each>
                    </xsl:when>
                    <xsl:when test="(d:GetTitle($myObj)='') and (d:GetYear($myObj)='')">
                        <author>
                            <xsl:text>(</xsl:text>
                            <!--Calling GetAuthor function to get the value of the Author-->
                            <xsl:value-of select="d:GetAuthor($myObj)"/>
                            <xsl:text>)</xsl:text>
                        </author>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:text>(</xsl:text>
                        <author>
                            <!--Calling GetAuthor function to get the value of the Author-->
                            <xsl:value-of select="d:GetAuthor($myObj)"/>
                        </author>
                        <xsl:text>,</xsl:text>
                        <!--Calling GetYear function to get the value of the Year-->
                        <xsl:value-of select="d:GetYear($myObj)"/>
                        <xsl:text>)</xsl:text>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <!--Tmplate to capture section information-->
    <xsl:template name="SectionInfo">
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:for-each select="following-sibling::*">
            <xsl:message terminate="no">progress:parahandler</xsl:message>
            <xsl:choose>
                <!--Checking for section break-->
                <xsl:when test="w:pPr/w:sectPr">
                    <xsl:if test="d:GetSectionPageStart($myObj)=1">
                        <xsl:sequence select="d:sink(d:SectionCounter($myObj,w:pPr/w:sectPr/w:pgNumType/@w:fmt,w:pPr/w:sectPr/w:pgNumType/@w:start))"/> <!-- empty -->
                    </xsl:if>
                </xsl:when>
                <!--Checking for section break-->
                <xsl:when test="name()='w:sectPr'">
                    <xsl:if test="d:GetSectionPageStart($myObj)=1">
                        <xsl:sequence select="d:sink(d:SectionCounter($myObj,w:pgNumType/@w:fmt,w:pgNumType/@w:start))"/> <!-- empty -->
                    </xsl:if>
                </xsl:when>
            </xsl:choose>
        </xsl:for-each>
        <xsl:sequence select="d:sink(d:InitalizeSectionPageStart($myObj))"/> <!-- empty -->
    </xsl:template>

    <!--Template to Implement Blocquote using Blocklist Template -->
    <xsl:template name="Blocklist">
        <xsl:param name="prmTrack"/>
        <xsl:param name="VERSION"/>
        <xsl:param name="custom"/>
        <xsl:param name="blkcharstyle"/>
        <xsl:if test="count(preceding-sibling::node()[1]/w:pPr/w:pStyle[substring(@w:val,1,5)='Block'])=0">
            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','blockquote ','&gt;')"/>
        </xsl:if>
        <xsl:choose>
            <!--Checking for 'Blockquote-AuthorDAISY' style-->
            <xsl:when test="w:pPr/w:pStyle[@w:val='Blockquote-AuthorDAISY']">
                <author>
                    <xsl:call-template name="ParaHandler">
                        <xsl:with-param name="flag" select="'0'"/>
                        <xsl:with-param name="custom" select="$custom"/>
                        <xsl:with-param name="charparahandlerStyle" select="$blkcharstyle"/>
                    </xsl:call-template>
                </author>
            </xsl:when>
            <!--Checking for List in BlockQuote Element-->
            <xsl:when test="((w:pPr/w:numPr/w:ilvl) and (w:pPr/w:numPr/w:numId))">
                <!--variable checkilvl holds level(w:ilvl) value of the List-->
                <xsl:call-template name="List">
                    <xsl:with-param name="verlist" select="$VERSION"/>
                    <xsl:with-param name="listcharStyle" select="$blkcharstyle"/>
                </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
                <xsl:if test="not(w:pPr/pStyle/@w:val='List-HeadingDAISY')">
                    <!--Calling ParagraphStyle Template-->
                    <xsl:call-template name="ParagraphStyle">
                        <xsl:with-param name="flag" select="'0'"/>
                        <xsl:with-param name="custom" select="$custom"/>
                        <xsl:with-param name="characterparaStyle" select="$blkcharstyle"/>
                    </xsl:call-template>
                </xsl:if>
            </xsl:otherwise>
        </xsl:choose>
        <xsl:if test="count(following-sibling::node()[1]/w:pPr/w:pStyle[substring(@w:val,1,5)='Block'])=0">
            <!--Closing blockquote element-->
            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/blockquote ','&gt;')"/>
        </xsl:if>
    </xsl:template>

    <!--Template for Custom styles-->
    <xsl:template name="CustomStyles">
        <xsl:param name="customTag"/>
        <xsl:param name="custom"/>
        <xsl:param name="txt"/>
        <xsl:param name="customcharStyle"/>
        <xsl:variable name="aquote">"</xsl:variable>
        <xsl:choose>
            <!--Checking for SampleDAISY/HTMLSample custom character style-->
            <xsl:when test="($customTag='SampleDAISY') or ($customTag='HTMLSample')" >
                <samp>
                    <xsl:attribute name="xml:space">preserve</xsl:attribute>
                    <xsl:call-template name="CustomCharStyle">
                        <xsl:with-param name="characterStyle" select="$customcharStyle"/>
                        <xsl:with-param name="txt" select="$txt"/>
                    </xsl:call-template>
                </samp>
            </xsl:when>
            <!--Checking for QuotationDAISY custom character style-->
            <xsl:when test="$customTag='QuotationDAISY'" >
                <q>
                    <xsl:call-template name="CustomCharStyle">
                        <xsl:with-param name="characterStyle" select="$customcharStyle"/>
                        <xsl:with-param name="txt" select="$txt"/>
                    </xsl:call-template>
                </q>
            </xsl:when>
            <!--Checking for CodeDAISY/HTMLCode custom character style-->
            <xsl:when test="($customTag='CodeDAISY') or ($customTag='HTMLCode')">
                <xsl:if test="count(preceding-sibling::w:r[1]/w:rPr/w:rStyle[contains(@w:val,'Code')])=0">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','code ','xml:space=',$aquote,'preserve',$aquote,'&gt;')"/>
                    <xsl:sequence select="d:sink(d:CodeFlag($myObj))"/> <!-- empty -->
                </xsl:if>
                <xsl:call-template name="CustomCharStyle">
                    <xsl:with-param name="characterStyle" select="$customcharStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                </xsl:call-template>
                <xsl:if test="count(following-sibling::w:r[1]/w:rPr/w:rStyle[contains(@w:val,'Code')])=0">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/code','&gt;')"/>
                    <xsl:sequence select="d:sink(d:InitializeCodeFlag($myObj))"/> <!-- empty -->
                </xsl:if>
            </xsl:when>
            <!--Checking for SentDAISY custom character style-->
            <xsl:when test="($customTag='SentDAISY')">
                <xsl:if test="count(preceding-sibling::w:r[1]/w:rPr/w:rStyle[contains(@w:val,'Sent')])=0">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','sent','&gt;')"/>
                </xsl:if>
                <xsl:call-template name="CustomCharStyle">
                    <xsl:with-param name="characterStyle" select="$customcharStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                </xsl:call-template>
                <xsl:if test="count(following-sibling::w:r[1]/w:rPr/w:rStyle[contains(@w:val,'Sent')])=0">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/sent','&gt;')"/>
                </xsl:if>
            </xsl:when>
            <!--Checking for SpanDAISY custom character style-->
            <xsl:when test="($customTag='SpanDAISY')">
                <xsl:if test="count(preceding-sibling::w:r[1]/w:rPr/w:rStyle[contains(@w:val,'Span')])=0">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','span','&gt;')"/>
                </xsl:if>
                <xsl:call-template name="CustomCharStyle">
                    <xsl:with-param name="characterStyle" select="$customcharStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                </xsl:call-template>
                <xsl:if test="count(following-sibling::w:r[1]/w:rPr/w:rStyle[contains(@w:val,'Span')])=0">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/span','&gt;')"/>
                </xsl:if>
            </xsl:when>
            <!--Checking for DefinitionDAISY/HTMLDefinition custom character style-->
            <xsl:when test="($customTag='DefinitionDAISY') or ($customTag='HTMLDefinition')">
                <dfn>
                    <xsl:call-template name="CustomCharStyle">
                        <xsl:with-param name="characterStyle" select="$customcharStyle"/>
                        <xsl:with-param name="txt" select="$txt"/>
                    </xsl:call-template>
                </dfn>
            </xsl:when>
            <!--Checking for CitationDAISY/HTMLCite custom character style-->
            <xsl:when test="($customTag='CitationDAISY')or ($customTag='HTMLCite')">
                <cite>
                    <xsl:call-template name="CustomCharStyle">
                        <xsl:with-param name="characterStyle" select="$customcharStyle"/>
                        <xsl:with-param name="txt" select="$txt"/>
                    </xsl:call-template>
                </cite>
            </xsl:when>
            <!--Checking for KeyboardInputDAISY/HTMLKeyboard custom character style-->
            <xsl:when test="($customTag='KeyboardInputDAISY') or ($customTag='HTMLKeyboard')">
                <kbd>
                    <xsl:call-template name="CustomCharStyle">
                        <xsl:with-param name="characterStyle" select="$customcharStyle"/>
                        <xsl:with-param name="txt" select="$txt"/>
                    </xsl:call-template>
                </kbd>
            </xsl:when>
            <!--Checking for Page number custom style-->
            <xsl:when test="($customTag='PageNumberDAISY') and ($custom='Custom')">
                <xsl:if test="count(preceding-sibling::w:r[1]/w:rPr/w:rStyle[@w:val='PageNumberDAISY'])=0">
                    <xsl:variable name="page">
                        <xsl:choose>
                            <xsl:when test="string(number(w:t))='NaN'">
                                <xsl:value-of select="'special'"/>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of select="'normal'"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:variable>
                    <!--Opening page number tag-->
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','pagenum ','page=',$aquote,$page,$aquote,' id=',$aquote,concat('page',d:GeneratePageId($myObj)),$aquote,'&gt;')"/>
                </xsl:if>
                <!--Calling template for page number text-->
                <xsl:call-template name="CustomCharStyle">
                    <xsl:with-param name="characterStyle" select="$customcharStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                </xsl:call-template>
                <xsl:if test="count(following-sibling::node()[1]/w:rPr/w:rStyle[@w:val='PageNumberDAISY'])=0">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/pagenum','&gt;')"/>
                </xsl:if>
            </xsl:when>
            <xsl:when test="($customTag='LineNumberDAISY')">
                <xsl:choose>
                    <xsl:when test="(../w:pPr/w:pStyle[@w:val='PoemDAISY']) or (../w:pPr/w:pStyle[@w:val='AddressDAISY'])">
                        <linenum>
                            <xsl:call-template name="CustomCharStyle">
                                <xsl:with-param name="characterStyle" select="$customcharStyle"/>
                                <xsl:with-param name="txt" select="$txt"/>
                            </xsl:call-template>
                        </linenum>
                    </xsl:when>
                    <xsl:when test="(../w:pPr/w:pStyle[@w:val='DefinitionDataDAISY'])">
                        <line>
                            <linenum>
                                <xsl:call-template name="CustomCharStyle">
                                    <xsl:with-param name="characterStyle" select="$customcharStyle"/>
                                    <xsl:with-param name="txt" select="$txt"/>
                                </xsl:call-template>
                            </linenum>
                        </line>
                    </xsl:when>
                    <xsl:when test="(d:Getlinenumflag($myObj)=1)">
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
                        <line>
                            <linenum>
                                <xsl:call-template name="CustomCharStyle">
                                    <xsl:with-param name="characterStyle" select="$customcharStyle"/>
                                    <xsl:with-param name="txt" select="$txt"/>
                                </xsl:call-template>
                            </linenum>
                        </line>
                        <xsl:choose>
                            <xsl:when test="(following-sibling::node()[1][name()='w:r']) and (not(following-sibling::node()[1]/w:rPr/w:rStyle[contains(@w:val,'Line')]))">
                                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p','&gt;')"/>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:sequence select="d:sink(d:Resetlinenumflag($myObj))"/> <!-- empty -->
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:when>
                </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
                <xsl:call-template name="CustomCharStyle">
                    <xsl:with-param name="characterStyle" select="$customcharStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                </xsl:call-template>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <!--Template for different inbuilt character styles-->
    <xsl:template name="CustomCharStyle">
        <xsl:param name="characterStyle"/>
        <xsl:param name="txt"/>
        <xsl:choose>
            <!--Strong and Emp-->
            <xsl:when test="(w:rPr/w:b or w:rPr/w:rStyle[@w:val='Strong']) and (w:rPr/w:i or w:rPr/w:rStyle[@w:val='Emphasis']) and not((w:rPr/w:vertAlign[@w:val ='superscript']) or (w:rPr/w:vertAlign[@w:val ='subscript']) or (w:rPr/w:rStyle[@w:val ='FootnoteReference']) or (w:rPr/w:rStyle[@w:val ='EndnoteReference']) or (w:footnoteReference) or (w:endnoteReference))">
                <strong>
                    <em>
                        <xsl:call-template name="CharacterStyles">
                            <xsl:with-param name="characterStyle" select="$characterStyle"/>
                            <xsl:with-param name="txt" select="$txt"/>
                        </xsl:call-template>
                    </em>
                </strong>
            </xsl:when>
            <!--Strong-->
            <xsl:when test="((w:rPr/w:b) or (w:rPr/w:rStyle[@w:val='Strong']))    and not((w:rPr/w:vertAlign[@w:val ='subscript']) or (w:rPr/w:vertAlign[@w:val ='superscript'])or (w:rPr/w:rStyle[@w:val ='FootnoteReference']) or (w:rPr/w:rStyle[@w:val ='EndnoteReference']) or (w:footnoteReference) or (w:endnoteReference))">
                <strong>
                    <xsl:call-template name="CharacterStyles">
                        <xsl:with-param name="characterStyle" select="$characterStyle"/>
                        <xsl:with-param name="txt" select="$txt"/>
                    </xsl:call-template>
                </strong>
            </xsl:when>
            <!--Emp-->
            <xsl:when test="((w:rPr/w:i) or (w:rPr/w:rStyle[@w:val='Emphasis']))    and not((w:rPr/w:vertAlign[@w:val ='superscript']) or (w:rPr/w:vertAlign[@w:val ='subscript'])or (w:rPr/w:rStyle[@w:val ='FootnoteReference']) or (w:rPr/w:rStyle[@w:val ='EndnoteReference']) or (w:footnoteReference) or (w:endnoteReference))">
                <em>
                    <xsl:call-template name="CharacterStyles">
                        <xsl:with-param name="characterStyle" select="$characterStyle"/>
                        <xsl:with-param name="txt" select="$txt"/>
                    </xsl:call-template>
                </em>
            </xsl:when>
            <!--Subscript Strong and Emp-->
            <xsl:when test="(w:rPr/w:vertAlign[@w:val ='subscript']) and (w:rPr/w:b and w:rPr/w:i) and not((w:rPr/w:rStyle[@w:val ='FootnoteReference']) or (w:rPr/w:rStyle[@w:val ='EndnoteReference']) or (w:footnoteReference) or (w:endnoteReference))">
                <sub>
                    <strong>
                        <em>
                            <xsl:call-template name="CharacterStyles">
                                <xsl:with-param name="characterStyle" select="$characterStyle"/>
                                <xsl:with-param name="txt" select="$txt"/>
                            </xsl:call-template>
                        </em>
                    </strong>
                </sub>
            </xsl:when>
            <!--Subscript and Strong-->
            <xsl:when test="(w:rPr/w:vertAlign[@w:val ='subscript']) and (w:rPr/w:b) and not((w:rPr/w:rStyle[@w:val ='FootnoteReference']) or (w:rPr/w:rStyle[@w:val ='EndnoteReference']) or (w:footnoteReference) or (w:endnoteReference))">
                <sub>
                    <strong>
                        <xsl:call-template name="CharacterStyles">
                            <xsl:with-param name="characterStyle" select="$characterStyle"/>
                            <xsl:with-param name="txt" select="$txt"/>
                        </xsl:call-template>
                    </strong>
                </sub>
            </xsl:when>
            <!--Subscript and Emp-->
            <xsl:when test="(w:rPr/w:vertAlign[@w:val ='subscript']) and (w:rPr/w:i) and not((w:rPr/w:rStyle[@w:val ='FootnoteReference']) or (w:rPr/w:rStyle[@w:val ='EndnoteReference']) or (w:footnoteReference) or (w:endnoteReference))">
                <sub>
                    <em>
                        <xsl:call-template name="CharacterStyles">
                            <xsl:with-param name="characterStyle" select="$characterStyle"/>
                            <xsl:with-param name="txt" select="$txt"/>
                        </xsl:call-template>
                    </em>
                </sub>
            </xsl:when>
            <!--Subscript-->
            <xsl:when test="w:rPr/w:vertAlign[@w:val ='subscript'] and not((w:rPr/w:rStyle[@w:val ='FootnoteReference']) or (w:rPr/w:rStyle[@w:val ='EndnoteReference'])    or (w:footnoteReference) or (w:endnoteReference))">
                <sub>
                    <xsl:call-template name="CharacterStyles">
                        <xsl:with-param name="characterStyle" select="$characterStyle"/>
                        <xsl:with-param name="txt" select="$txt"/>
                    </xsl:call-template>
                </sub>
            </xsl:when>
            <!--Superscript Strong and Emp-->
            <xsl:when test="(w:rPr/w:vertAlign[@w:val ='superscript']) and (w:rPr/w:b and w:rPr/w:i) and not((w:rPr/w:rStyle[@w:val ='FootnoteReference']) or (w:rPr/w:rStyle[@w:val ='EndnoteReference'])    or (w:footnoteReference) or (w:endnoteReference))">
                <sup>
                    <strong>
                        <em>
                            <xsl:call-template name="CharacterStyles">
                                <xsl:with-param name="characterStyle" select="$characterStyle"/>
                                <xsl:with-param name="txt" select="$txt"/>
                            </xsl:call-template>
                        </em>
                    </strong>
                </sup>
            </xsl:when>
            <!--Superscript and Strong-->
            <xsl:when test="(w:rPr/w:vertAlign[@w:val ='superscript']) and (w:rPr/w:b) and not((w:rPr/w:rStyle[@w:val ='FootnoteReference']) or (w:rPr/w:rStyle[@w:val ='EndnoteReference']) or (w:footnoteReference) or (w:endnoteReference))">
                <sup>
                    <strong>
                        <xsl:call-template name="CharacterStyles">
                            <xsl:with-param name="characterStyle" select="$characterStyle"/>
                            <xsl:with-param name="txt" select="$txt"/>
                        </xsl:call-template>
                    </strong>
                </sup>
            </xsl:when>
            <!--Superscript and Emp-->
            <xsl:when test="(w:rPr/w:vertAlign[@w:val ='superscript']) and (w:rPr/w:i) and not((w:rPr/w:rStyle[@w:val ='FootnoteReference']) or (w:rPr/w:rStyle[@w:val ='EndnoteReference']) or (w:footnoteReference) or (w:endnoteReference))">
                <sup>
                    <em>
                        <xsl:call-template name="CharacterStyles">
                            <xsl:with-param name="characterStyle" select="$characterStyle"/>
                            <xsl:with-param name="txt" select="$txt"/>
                        </xsl:call-template>
                    </em>
                </sup>
            </xsl:when>
            <!--Superscript-->
            <xsl:when test="w:rPr/w:vertAlign[@w:val ='superscript'] and not((w:rPr/w:rStyle[@w:val ='FootnoteReference']) or (w:rPr/w:rStyle[@w:val ='EndnoteReference']) or (w:footnoteReference) or (w:endnoteReference))">
                <sup>
                    <xsl:call-template name="CharacterStyles">
                        <xsl:with-param name="characterStyle" select="$characterStyle"/>
                        <xsl:with-param name="txt" select="$txt"/>
                    </xsl:call-template>
                </sup>
            </xsl:when>
            <!--Checking for WordDAISY custom character style-->
            <xsl:when test="(w:rPr/w:rStyle[contains(@w:val,'Word')])">
                <w>
                    <xsl:call-template name="CharacterStyles">
                        <xsl:with-param name="characterStyle" select="$characterStyle"/>
                        <xsl:with-param name="txt" select="$txt"/>
                    </xsl:call-template>
                </w>
            </xsl:when>
            <xsl:when test="w:instrText">
                <xsl:if test="not(preceding-sibling::w:r[1]/w:fldChar[@w:fldCharType='begin'])">
                    <xsl:value-of select="w:instrText"/>
                </xsl:if>
            </xsl:when>

            <!--Checking conditions for Character styles-->

            <xsl:when test="(w:t) and (not(w:rPr/w:rStyle[@w:val='DefinitionTermDAISY']))">
                <!--Hyphen-->
                <xsl:if test="w:noBreakHyphen">
                    <xsl:text>-</xsl:text>
                </xsl:if>
                <xsl:call-template name="CharacterStyles">
                    <xsl:with-param name="characterStyle" select="$characterStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                </xsl:call-template>
            </xsl:when>

            <!--Fidility Loss Report-->
            <xsl:otherwise>
                <xsl:for-each select="./node()">
                    <xsl:call-template name="CharacterStyles">
                        <xsl:with-param name="characterStyle" select="$characterStyle"/>
                        <xsl:with-param name="txt" select="$txt"/>
                    </xsl:call-template>
                    <xsl:message terminate="no">progress:parahandler</xsl:message>
                    <xsl:choose>
                        <xsl:when test="name()='w:commentReference'">
                            <!--Capturing fidility loss-->
                            <xsl:message terminate="no">translation.oox2Daisy.commentReference</xsl:message>
                            <!--Capturing fidility loss-->
                        </xsl:when>
                        <xsl:when test="name()='w:object'">
                            <xsl:message terminate="no">translation.oox2Daisy.object</xsl:message>
                        </xsl:when>
                        <xsl:otherwise>
                            <!--Capturing fidility loss-->
                            <xsl:if test="not((name()='w:rPr')or(name()='w:fldSimple') or (name()='w:lastRenderedPageBreak') or (name()='w:br') or (name()='w:tab')or(name()='w:fldChar') or (name()='w:t'))">
                                <xsl:message terminate="no">
                                    <xsl:value-of select="concat('translation.oox2Daisy.UncoveredElement|',name())"/>
                                </xsl:message>
                            </xsl:if>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:for-each>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template name="List">
        <xsl:param name="verlist"/>
        <xsl:param name="prmTrack"/>
        <xsl:param name="custom"/>
        <xsl:param name="sOperators"/>
        <xsl:param name="sMinuses"/>
        <xsl:param name="sNumbers"/>
        <xsl:param name="sZeros"/>
        <xsl:param name="listcharStyle"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <!--variable checkilvl holds level(w:ilvl) value of the List-->
        <!--NOTE:Use GetCheckLvlInt function that return 0 if node is not exists-->
        <xsl:variable name="checkilvl" select="d:GetCheckLvlInt($myObj,w:pPr/w:numPr/w:ilvl/@w:val)"/>
        <!--Variable checknumId holds the w:numId value which specifies the type of numbering in the list-->
        <xsl:variable name="checknumId" select="w:pPr/w:numPr/w:numId/@w:val"/>

        <xsl:variable name="CheckNumId" select="d:CheckNumID($myObj,$checknumId)"/>
        <xsl:if test="$CheckNumId ='True'">
            <xsl:sequence select="d:sink(d:StartNewListCounter($myObj,$checknumId))"/> <!-- empty -->
        </xsl:if>

        <xsl:if test="string-length(preceding-sibling::node()[1]/w:pPr/w:numPr/w:ilvl/@w:val)=0 or preceding-sibling::w:p[1]/w:pPr/w:rPr/w:vanish" >
            <xsl:if test="($checkilvl &gt; 0)">
                <xsl:call-template name="recursive">
                    <xsl:with-param name="rec" select="$checkilvl"/>
                </xsl:call-template>
                <xsl:sequence select="d:Increment($myObj,$checkilvl)"/> <!-- empty -->
            </xsl:if>
        </xsl:if>

        <xsl:if test="string-length(preceding-sibling::node()[1]/w:pPr/w:numPr/w:ilvl/@w:val) = 0
                                                                        or (preceding-sibling::node()[1]/w:pPr/w:numPr/w:ilvl/@w:val &lt; $checkilvl
                                                                                        and not(preceding-sibling::node()[1]/w:pPr/w:numPr/w:ilvl/@w:val=$checkilvl))
                                                                        or preceding-sibling::node()[1]/w:pPr/w:pStyle[substring(@w:val,1,7)='Heading']
                                                                        or preceding-sibling::node()[1]/w:pPr/w:rPr/w:vanish">
            <xsl:variable name="val">
                <xsl:value-of select="$numberingXml//w:numbering/w:num[@w:numId=$checknumId]/w:abstractNumId/@w:val"/>
            </xsl:variable>
            <!--Checking numbering.xml for the type of List-->
            <xsl:variable    name="type">
                <xsl:value-of select="$numberingXml//w:numbering/w:abstractNum[@w:abstractNumId=$val]/w:lvl[@w:ilvl=$checkilvl]/w:numFmt/@w:val"/>
            </xsl:variable>
            <!--Checking for Ordered List type-->
            <xsl:choose>
                <xsl:when test="$type='decimal'">
                    <xsl:variable name="aquote">"</xsl:variable>
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;', 'list', ' type=', $aquote, 'ol', $aquote,' start=', $aquote, d:GetListCounter($myObj,$checkilvl, $checknumId), $aquote, '&gt;')"/>
                </xsl:when>
                <xsl:when test="($type='lowerLetter') or ($type='lowerRoman') or ($type='upperRoman') or ($type='upperLetter')or ($type='decimalZero')">
                    <xsl:variable name="aquote">"</xsl:variable>
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','list ','type=',$aquote,'pl',$aquote,'&gt;')"/>
                </xsl:when>
                <!--Checking for Unordered list and Preformatted List type-->
                <xsl:when test="$type='bullet'">
                    <xsl:choose>
                        <xsl:when test ="$numberingXml//w:numbering/w:abstractNum[@w:abstractNumId=$val]/w:lvl[@w:ilvl=$checkilvl]/w:lvlPicBulletId">
                            <xsl:variable name="aquote">"</xsl:variable>
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','list ','type=',$aquote,'pl',$aquote,'&gt;')"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:variable name="aquote">"</xsl:variable>
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','list ','type=',$aquote,'ul',$aquote,'&gt;')"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:when>
                <xsl:when test="$type='none'">
                    <xsl:variable name="aquote">"</xsl:variable>
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','list ','type=',$aquote,'pl',$aquote,'&gt;')"/>
                </xsl:when>
                <!--If in word/numbering.xml,numbering formatis having any style(means no w:numFmt element), we are taking default style as ordered list-->
                <xsl:when test="$type=''">
                    <xsl:variable name="aquote">"</xsl:variable>
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','list ','type=',$aquote,'ol',$aquote,'&gt;')"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:variable name="aquote">"</xsl:variable>
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','list ','type=',$aquote,'pl',$aquote,'&gt;')"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:if>
        <xsl:if test="d:ListHeadingFlag($myObj)&gt;0">
            <xsl:value-of disable-output-escaping="yes" select="d:PeekListHeading($myObj)"/>
            <xsl:sequence select="d:sink(d:ReSetListHeadingFlag($myObj))"/> <!-- empty -->
        </xsl:if>

        <!--Checking the current level with the next level-->
        <xsl:variable name="PeekLevel" select="d:ListPeekLevel($myObj)"/>
        <xsl:if test="d:Difference($checkilvl,$PeekLevel)">
            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/li','&gt;')"/>
            <xsl:call-template name="ComplexListClose">
                <xsl:with-param name="close" select="$checkilvl"/>
            </xsl:call-template>
        </xsl:if>

        <!--Closing the List-->
        <xsl:call-template name="closelist">
            <xsl:with-param name="close" select="$checkilvl"/>
        </xsl:call-template>

        <xsl:variable name="val">
            <xsl:value-of select="$numberingXml//w:numbering/w:num[@w:numId=$checknumId]/w:abstractNumId/@w:val"/>
        </xsl:variable>
        <!--Checking numbering.xml for the type of List-->
        <xsl:variable    name="numFormat">
            <xsl:value-of select="$numberingXml//w:numbering/w:abstractNum[@w:abstractNumId=$val]/w:lvl[@w:ilvl=$checkilvl]/w:numFmt/@w:val"/>
        </xsl:variable>
        <xsl:variable    name="lvlText">
            <xsl:value-of select="$numberingXml//w:numbering/w:abstractNum[@w:abstractNumId=$val]/w:lvl[@w:ilvl=$checkilvl]/w:lvlText/@w:val"/>
        </xsl:variable>

        <xsl:call-template name="recStart">
            <xsl:with-param name="abstLevel" select="$val"/>
            <xsl:with-param name="level" select="$checkilvl"/>
        </xsl:call-template>

        <!--Opening the List-->
        <xsl:call-template name="addlist">
            <xsl:with-param name="openId" select="$checknumId"/>
            <xsl:with-param name="openlvl" select="$checkilvl"/>
            <xsl:with-param name="verlist" select="$verlist"/>
            <xsl:with-param name="custom" select="$custom"/>
            <xsl:with-param name="numFmt" select="$numFormat"/>
            <xsl:with-param name="lText" select="$lvlText"/>
            <xsl:with-param name="lstcharStyle" select="$listcharStyle"/>
        </xsl:call-template>

        <!--Closing the current List item(li) element when there is no nested list-->
        <xsl:variable name="LPeekLevel" select="d:ListPeekLevel($myObj)"/>
        <xsl:if test="($LPeekLevel = $checkilvl
                                                                                and following-sibling::node()[1][w:pPr/w:numPr/w:ilvl/@w:val = $checkilvl]
                                                                                and not(following-sibling::node()[1]/w:pPr/w:pStyle[substring(@w:val,1,7)='Heading'])
                                                                                and not(following-sibling::node()[1]/w:pPr/w:rPr/w:vanish))
                                                                ">
            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/li','&gt;')"/>
        </xsl:if>

        <!--Insert page break in list-->
        <xsl:call-template name="PageInList">
            <xsl:with-param name="custom" select="$custom"/>
        </xsl:call-template>

        <!--Closing all the nested Lists-->
        <xsl:if test="(count(following-sibling::node()[1][w:pPr/w:numPr/w:ilvl/@w:val])=0
                                                                        or following-sibling::w:p[1]/w:pPr/w:pStyle[substring(@w:val,1,7)='Heading']
                                                                        or (following-sibling::w:p[1]/w:pPr/w:rPr/w:vanish))
                                                                ">
            <xsl:call-template name="CloseLastlist">
                <xsl:with-param name="close" select="0"/>
                <xsl:with-param name="custom" select="$custom"/>
            </xsl:call-template>
        </xsl:if>

    </xsl:template>

    <!--Template for implementing paragraph styles-->
    <xsl:template name="StyleContainer">
        <xsl:param name="prmTrack"/>
        <xsl:param name="VERSION"/>
        <xsl:param name="custom"/>
        <xsl:param name="styleHeading"/>
        <xsl:param name="sOperators"/>
        <xsl:param name="sMinuses"/>
        <xsl:param name="sNumbers"/>
        <xsl:param name="sZeros"/>
        <xsl:param name="mastersubstyle"/>
        <xsl:param name="txt"/>
        <xsl:param name="imgOptionStyle"/>
        <xsl:param name="dpiStyle"/>
        <xsl:param name="characterStyle"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:choose>
            <!--Checking for List in Blockquote-->
            <xsl:when test="w:pPr/w:pStyle[substring(@w:val,1,5)='Block']">
                <!--Calling Blocklist Template-->
                <xsl:call-template name="Blocklist">
                    <xsl:with-param name="prmTrack" select="$prmTrack"/>
                    <xsl:with-param name="VERSION" select="$VERSION"/>
                    <xsl:with-param name="custom" select="$custom"/>
                    <xsl:with-param name="blkcharstyle" select="$characterStyle"/>
                </xsl:call-template>
            </xsl:when>

            <xsl:when test="w:pPr/w:pStyle[substring(@w:val,1,3)='TOC'] and not(preceding::w:pPr/w:pStyle[substring(@w:val,1,3)='TOC'])">
                <!--Save Level before closing all levels-->
                <xsl:variable name="PeekLevel" select="d:PeekLevel($myObj)"/>
                <!--Close all levels before Table Of Contents-->
                <xsl:call-template name="CloseLevel">
                    <xsl:with-param name="CurrentLevel" select="1"/>
                </xsl:call-template>
                <!--Calling Template to add Table Of Contents-->
                <xsl:call-template name="TableOfContents">
                    <xsl:with-param name="custom" select="$custom"/>
                </xsl:call-template>
                <!--Open $PeekLevel levels after Table Of Contents-->
                <xsl:call-template name="AddLevel">
                    <xsl:with-param name="levelValue" select="$PeekLevel"/>
                    <xsl:with-param name="check" select="1"/>
                    <xsl:with-param name="verhead" select="$VERSION"/>
                    <xsl:with-param name="custom" select="$custom"/>
                    <xsl:with-param name="sOperators" select="$sOperators"/>
                    <xsl:with-param name="sMinuses" select="$sMinuses"/>
                    <xsl:with-param name="sNumbers" select="$sNumbers"/>
                    <xsl:with-param name="sZeros" select="$sZeros"/>
                    <xsl:with-param name="mastersubhead" select="$mastersubstyle"/>
                    <xsl:with-param name="txt" select="0"/>
                    <xsl:with-param name="lvlcharStyle" select="$characterStyle"/>
                </xsl:call-template>
            </xsl:when>

            <xsl:when test="w:pPr/w:pStyle[substring(@w:val,1,7)='Heading' or @w:val/d:CompareHeading(.,$styleHeading)=1] and (not(name(..)='w:tc'))">
                <!--calling Close level template for closing all the higher levels-->
                <xsl:if test="(w:r/w:lastRenderedPageBreak) or (w:r/w:br/@w:type='page')">
                    <xsl:call-template name="footnote">
                        <xsl:with-param name="verfoot" select="$VERSION"/>
                        <xsl:with-param name="mastersubfoot" select="$mastersubstyle"/>
                        <xsl:with-param name="sOperators" select="$sOperators"/>
                        <xsl:with-param name="sMinuses" select="$sMinuses"/>
                        <xsl:with-param name="sNumbers" select="$sNumbers"/>
                        <xsl:with-param name="sZeros" select="$sZeros"/>
                    </xsl:call-template>
                </xsl:if>

                <xsl:choose>
                    <xsl:when test="((w:pPr/w:numPr/w:ilvl) and (w:pPr/w:numPr/w:numId))">
                        <xsl:variable name="text_heading">
                            <!--variable checkilvl holds level(w:ilvl) value of the List-->
                            <!--NOTE:Use GetCheckLvlInt function that return 0 if node is not exists-->
                            <xsl:variable name="checkilvl" select="d:GetCheckLvlInt($myObj,w:pPr/w:numPr/w:ilvl/@w:val)"/>
                            <!--Variable checknumId holds the w:numId value which specifies the type of numbering in the list-->
                            <xsl:variable name="checknumId" select="w:pPr/w:numPr/w:numId/@w:val"/>
                            <xsl:call-template name="HeadingsPart">
                                <xsl:with-param name="checkilvl" select="$checkilvl"/>
                                <xsl:with-param name="checknumId" select="$checknumId"/>
                                <xsl:with-param name="doc" select="'Document'"/>
                            </xsl:call-template>
                            <xsl:value-of select="d:RetrieveHeadingPart($myObj)"/>
                        </xsl:variable>

                        <xsl:variable name="absValue">
                            <!--NOTE:Use GetCheckLvlInt function that return 0 if node is not exists-->
                            <xsl:sequence select="d:sink(d:GetCheckLvlInt($myObj,w:pPr/w:numPr/w:ilvl/@w:val))"/> <!-- empty -->
                            <!--Variable checknumId holds the w:numId value which specifies the type of numbering in the list-->
                            <xsl:variable name="checknumId" select="w:pPr/w:numPr/w:numId/@w:val"/>
                            <xsl:value-of select="$numberingXml//w:numbering/w:num[@w:numId=$checknumId]/w:abstractNumId/@w:val"/>
                        </xsl:variable>

                        <xsl:call-template name="CloseLevel">
                            <xsl:with-param name="CurrentLevel" select="number(substring(w:pPr/w:pStyle/@w:val,string-length(w:pPr/w:pStyle/@w:val)))"/>
                        </xsl:call-template>

                        <!--calling AddLevel template for adding the levels-->
                        <xsl:call-template name="AddLevel">
                            <!--Passing parameter levelValue which holds the Heading type value-->
                            <xsl:with-param name="levelValue" select="w:pPr/w:numPr/w:ilvl/@w:val"/>
                            <xsl:with-param name="check" select="1"/>
                            <xsl:with-param name="verhead" select="$VERSION"/>
                            <xsl:with-param name="custom" select="$custom"/>
                            <xsl:with-param name="sOperators" select="$sOperators"/>
                            <xsl:with-param name="sMinuses" select="$sMinuses"/>
                            <xsl:with-param name="sNumbers" select="$sNumbers"/>
                            <xsl:with-param name="sZeros" select="$sZeros"/>
                            <xsl:with-param name="mastersubhead" select="$mastersubstyle"/>
                            <xsl:with-param name="abValue" select="$absValue"/>
                            <xsl:with-param name="txt" select="concat($text_heading,'!',w:pPr/w:numPr/w:numId/@w:val)"/>
                            <xsl:with-param name="lvlcharStyle" select="$characterStyle"/>
                        </xsl:call-template>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:variable name="text_heading">
                            <xsl:variable name="nameHeading" select="w:pPr/w:pStyle/@w:val"/>
                            <xsl:for-each select="$stylesXml//w:styles/w:style[@w:styleId=$nameHeading]">
                                <xsl:message terminate="no">progress:parahandler</xsl:message>
                                <xsl:choose>
                                    <xsl:when test="(./w:pPr/w:outlineLvl) and (./w:pPr/w:numPr/w:numId)">
                                        <!--NOTE:Use GetCheckLvlInt function that return 0 if node is not exists-->
                                        <xsl:variable name="checkilvl" select="d:GetCheckLvlInt($myObj,./w:pPr/w:outlineLvl/@w:val)"/>
                                        <xsl:variable name="checknumId" select="./w:pPr/w:numPr/w:numId/@w:val"/>
                                        <xsl:call-template name="HeadingsPart">
                                            <xsl:with-param name="checkilvl" select="$checkilvl"/>
                                            <xsl:with-param name="checknumId" select="$checknumId"/>
                                            <xsl:with-param name="doc" select="'Style'"/>
                                        </xsl:call-template>
                                        <xsl:value-of select="d:RetrieveHeadingPart($myObj)"/>
                                    </xsl:when>
                                    <xsl:when test="string-length(./w:pPr/w:numPr/w:numId)=0">
                                        <!--NOTE:Use GetCheckLvlInt function that return 0 if node is not exists-->
                                        <xsl:variable name="checkilvl" select="d:GetCheckLvlInt($myObj,./w:pPr/w:outlineLvl/@w:val)"/>
                                        <xsl:sequence select="d:sink(d:AddCurrHeadId($myObj,''))"/> <!-- empty -->
                                        <xsl:sequence select="d:sink(d:AddCurrHeadLevel($myObj,$checkilvl,'Style',''))"/> <!-- empty -->
                                        <xsl:value-of select="concat('','|',$checkilvl,'!','')"/>
                                    </xsl:when>
                                </xsl:choose>
                            </xsl:for-each>
                        </xsl:variable>

                        <xsl:variable name="absValue">
                            <xsl:variable name="nameHeading" select="w:pPr/w:pStyle/@w:val"/>
                            <xsl:for-each select="$stylesXml//w:styles/w:style[@w:styleId=$nameHeading]">
                                <xsl:message terminate="no">progress:parahandler</xsl:message>
                                <xsl:if test="(./w:pPr/w:outlineLvl) and (./w:pPr/w:numPr/w:numId)">
                                    <!--NOTE:Use GetCheckLvlInt function that return 0 if node is not exists-->
                                    <xsl:sequence select="d:sink(d:GetCheckLvlInt($myObj,./w:pPr/w:outlineLvl/@w:val))"/> <!-- empty -->
                                    <xsl:variable name="checknumId" select="./w:pPr/w:numPr/w:numId/@w:val"/>

                                    <xsl:value-of select="$numberingXml//w:numbering/w:num[@w:numId=$checknumId]/w:abstractNumId/@w:val"/>
                                </xsl:if>
                            </xsl:for-each>
                        </xsl:variable>


                        <xsl:variable name="ilvl">
                            <xsl:variable name="nameHeading" select="w:pPr/w:pStyle/@w:val"/>
                            <xsl:for-each select="$stylesXml//w:styles/w:style[@w:styleId=$nameHeading]">
                                <xsl:message terminate="no">progress:parahandler</xsl:message>
                                <xsl:choose>
                                    <xsl:when test="(./w:pPr/w:outlineLvl) and (./w:pPr/w:numPr/w:numId)">
                                        <xsl:value-of    select="./w:pPr/w:outlineLvl/@w:val"/>
                                    </xsl:when>
                                    <xsl:when test="string-length(./w:pPr/w:numPr/w:numId)=0">
                                        <xsl:value-of    select="./w:pPr/w:outlineLvl/@w:val"/>
                                    </xsl:when>
                                </xsl:choose>
                            </xsl:for-each>
                        </xsl:variable>

                        <xsl:call-template name="CloseLevel">
                            <xsl:with-param name="CurrentLevel" select="number(substring(w:pPr/w:pStyle/@w:val,string-length(w:pPr/w:pStyle/@w:val)))"/>
                        </xsl:call-template>

                        <xsl:call-template name="AddLevel">
                            <xsl:with-param name="levelValue" select="$ilvl[not(.='')]"/>
                            <xsl:with-param name="check" select="1"/>
                            <xsl:with-param name="verhead" select="$VERSION"/>
                            <xsl:with-param name="custom" select="$custom"/>
                            <xsl:with-param name="sOperators" select="$sOperators"/>
                            <xsl:with-param name="sMinuses" select="$sMinuses"/>
                            <xsl:with-param name="sNumbers" select="$sNumbers"/>
                            <xsl:with-param name="sZeros" select="$sZeros"/>
                            <xsl:with-param name="mastersubhead" select="$mastersubstyle"/>
                            <xsl:with-param name="abValue" select="$absValue"/>
                            <xsl:with-param name="txt" select="$text_heading"/>
                            <xsl:with-param name="lvlcharStyle" select="$characterStyle"/>
                        </xsl:call-template>

                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>

            <!--Checking for Bridgehead custom style-->
            <xsl:when test="(w:pPr/w:pStyle/@w:val='BridgeheadDAISY') and (not(name(..)='w:tc'))">
                <bridgehead>
                    <xsl:call-template name="ParaHandler">
                        <xsl:with-param name="flag" select="'0'"/>
                        <xsl:with-param name="VERSION" select="$VERSION"/>
                        <xsl:with-param name="custom" select="$custom"/>
                        <xsl:with-param name="imgOptionPara" select="$imgOptionStyle"/>
                        <xsl:with-param name="dpiPara" select="$dpiStyle"/>
                        <xsl:with-param name="txt" select="$txt"/>
                        <xsl:with-param name="charparahandlerStyle" select="$characterStyle"/>
                    </xsl:call-template>
                </bridgehead>
            </xsl:when>
            <xsl:otherwise>
                <xsl:choose>
                    <!--Cheaking for list heading-->
                    <xsl:when test="(w:pPr/w:pStyle/@w:val='List-HeadingDAISY')">
                        <xsl:variable name="tempListHeading">
                            <xsl:text disable-output-escaping="yes">&lt;hd&gt;</xsl:text>
                            <xsl:call-template name="ParaHandler">
                                <xsl:with-param name="flag" select="'0'"/>
                                <xsl:with-param name="VERSION" select="$VERSION"/>
                                <xsl:with-param name="custom" select="$custom"/>
                                <xsl:with-param name="imgOptionPara" select="$imgOptionStyle"/>
                                <xsl:with-param name="dpiPara" select="$dpiStyle"/>
                                <xsl:with-param name="charparahandlerStyle" select="$characterStyle"/>
                                <xsl:with-param name="txt" select="$txt"/>
                            </xsl:call-template>
                            <xsl:text disable-output-escaping="yes">&lt;/hd&gt;</xsl:text>
                        </xsl:variable>
                        <xsl:sequence select="d:sink(d:PushListHeading($myObj,$tempListHeading))"/> <!-- empty -->
                        <xsl:sequence select="d:sink(d:SetListHeadingFlag($myObj))"/> <!-- empty -->
                    </xsl:when>


                    <!--calling list template for implementing lists-->
                    <xsl:when test="(w:pPr/w:rPr/w:vanish) and (w:pPr/w:numPr/w:ilvl) and (w:pPr/w:numPr/w:numId)">
                        <xsl:variable name="checkCounter" select="d:IsList($myObj,w:pPr/w:numPr/w:numId/@w:val)"/>
                        <xsl:choose>
                            <xsl:when test="$checkCounter='ListTrue'">
                                <xsl:sequence select="d:IncrementListCounters($myObj,w:pPr/w:numPr/w:ilvl/@w:val,w:pPr/w:numPr/w:numId/@w:val)"/> <!-- empty -->
                            </xsl:when>
                            <xsl:otherwise>
                                <!--variable checkilvl holds level(w:ilvl) value of the List-->
                                <!--NOTE:Use GetCheckLvlInt function that return 0 if node is not exists-->
                                <xsl:variable name="checkilvl" select="d:GetCheckLvlInt($myObj,w:pPr/w:numPr/w:ilvl/@w:val)"/>
                                <!--Variable checknumId holds the w:numId value which specifies the type of numbering in the list-->
                                <xsl:variable name="checknumId" select="w:pPr/w:numPr/w:numId/@w:val"/>

                                <xsl:call-template name="VanishTemplate">
                                    <xsl:with-param name="checkilvl" select="$checkilvl"/>
                                    <xsl:with-param name="checknumId" select="$checknumId"/>
                                    <xsl:with-param name="checkCounter" select="$checkCounter"/>
                                </xsl:call-template>

                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:when>

                    <xsl:when test="((w:pPr/w:numPr/w:ilvl) and (w:pPr/w:numPr/w:numId) and not(w:pPr/w:rPr/w:vanish)) and not(w:pPr/w:pStyle[substring(@w:val,1,7)='Heading' or @w:val/d:CompareHeading(.,$styleHeading)=1])">
                        <xsl:call-template name="List">
                            <xsl:with-param name="prmTrack" select="$prmTrack"/>
                            <xsl:with-param name="custom" select="$custom"/>
                            <xsl:with-param name="listcharStyle" select="$characterStyle"/>
                        </xsl:call-template>
                    </xsl:when>

                    <xsl:otherwise>
                        <!--calling template named ParagraphStyle-->
                        <xsl:variable name="checkImageposition" select="d:GetCaptionsProdnotes($myObj)"/>
                        <xsl:if test="not((preceding-sibling::node()[$checkImageposition]/w:r/w:drawing)
                                                                                                                        or (preceding-sibling::node()[$checkImageposition]/w:r/w:pict)
                                                                                                                        or (preceding-sibling::node()[$checkImageposition]/w:r/w:object)
                                                                                                                        or (((w:pPr/w:pStyle/@w:val='Table-CaptionDAISY')
                                                                                                                                        or (w:pPr/w:pStyle/@w:val='Caption')
                                                                                                                                        or (./node()[name()='w:fldSimple']))
                                                                                                                        and ((preceding-sibling::node()[1][name()='w:tbl'])
                                                                                                                                        or (following-sibling::node()[1][name()='w:tbl']))))">
                            <xsl:call-template name="ParagraphStyle">
                                <xsl:with-param name="prmTrack" select="$prmTrack"/>
                                <xsl:with-param name="VERSION" select="$VERSION"/>
                                <xsl:with-param name="flag" select="'0'"/>
                                <xsl:with-param name="custom" select="$custom"/>
                                <xsl:with-param name="masterparastyle" select="$mastersubstyle"/>
                                <xsl:with-param name="imgOptionPara" select="$imgOptionStyle"/>
                                <xsl:with-param name="dpiPara" select="$dpiStyle"/>
                                <xsl:with-param name="txt" select="$txt"/>
                                <xsl:with-param name="characterparaStyle" select="$characterStyle"/>
                            </xsl:call-template>
                        </xsl:if>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template name="VanishTemplate">
        <xsl:param name="checkilvl"/>
        <xsl:param name="checknumId"/>
        <xsl:param name="checkCounter"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:variable name="val">
            <xsl:value-of select="$numberingXml//w:numbering/w:num[@w:numId=$checknumId]/w:abstractNumId/@w:val"/>
        </xsl:variable>
        <xsl:variable    name="lStartOverride">
            <xsl:value-of select="$numberingXml//w:numbering/w:num[@w:numId=$checknumId]/w:lvlOverride[@w:ilvl=$checkilvl]/w:startOverride/@w:val"/>
        </xsl:variable>
        <xsl:variable    name="lStart">
            <xsl:value-of select="$numberingXml//w:numbering/w:abstractNum[@w:abstractNumId=$val]/w:lvl[@w:ilvl=$checkilvl]/w:start/@w:val"/>
        </xsl:variable>

        <xsl:sequence select="d:sink(d:AddCurrHeadId($myObj,$checknumId))"/> <!-- empty -->
        <xsl:variable name="addCurrLvl" select="d:AddCurrHeadLevel($myObj,$checkilvl,'Vanish',$val)"/>

        <xsl:choose>
            <xsl:when test="$checkCounter='HeadTrue'">
                <xsl:choose>
                    <xsl:when test="string-length(substring-before($addCurrLvl,'|'))=0">
                        <xsl:choose>
                            <xsl:when test="not($lStartOverride='')">
                                <xsl:sequence select="d:sink(d:StartHeadingValueCtr($myObj,$checknumId,$val))"/> <!-- empty -->
                                <xsl:sequence select="d:sink(d:StartHeadingString($myObj,$checkilvl,$lStartOverride,$checknumId,$val,'Vanish','Yes'))"/> <!-- empty -->
                                <xsl:sequence select="d:sink(d:CopyToCurrCounter($myObj,$checknumId))"/> <!-- empty -->
                                <xsl:sequence select="d:IncrementHeadingCounters($myObj,w:pPr/w:numPr/w:ilvl/@w:val,$checknumId,$val)"/> <!-- empty -->
                                <xsl:sequence select="d:sink(d:CopyToBaseCounter($myObj,$checknumId))"/> <!-- empty -->
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:sequence select="d:sink(d:StartHeadingValueCtr($myObj,$checknumId,$val))"/> <!-- empty -->
                                <xsl:sequence select="d:sink(d:StartHeadingString($myObj,$checkilvl,$lStart,$checknumId,$val,'Vanish','No'))"/> <!-- empty -->
                                <xsl:sequence select="d:sink(d:CopyToCurrCounter($myObj,$checknumId))"/> <!-- empty -->
                                <xsl:sequence select="d:IncrementHeadingCounters($myObj,w:pPr/w:numPr/w:ilvl/@w:val,$checknumId,$val)"/> <!-- empty -->
                                <xsl:sequence select="d:sink(d:CopyToBaseCounter($myObj,$checknumId))"/> <!-- empty -->
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:when>
                </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
                <xsl:choose>
                    <xsl:when test="string-length(substring-before($addCurrLvl,'|'))=0">
                        <xsl:variable name="CheckNumId" select="d:CheckHeadingNumID($myObj,$checknumId)"/>
                        <xsl:if test="$CheckNumId ='True'">
                            <xsl:sequence select="d:sink(d:StartHeadingValueCtr($myObj,$checknumId,$val))"/> <!-- empty -->
                            <xsl:sequence select="d:sink(d:StartNewHeadingCounter($myObj,$checknumId,$val))"/> <!-- empty -->
                        </xsl:if>
                        <xsl:choose>
                            <xsl:when test="not($lStartOverride='')">
                                <xsl:sequence select="d:sink(d:CopyToCurrCounter($myObj,$checknumId))"/> <!-- empty -->
                                <xsl:sequence select="d:sink(d:StartHeadingValueCtr($myObj,$checknumId,$val))"/> <!-- empty -->
                                <xsl:sequence select="d:sink(d:StartHeadingString($myObj,$checkilvl,$lStartOverride,$checkCounter,$val,'Vanish','Yes'))"/> <!-- empty -->
                                <xsl:sequence select="d:IncrementHeadingCounters($myObj,w:pPr/w:numPr/w:ilvl/@w:val,$checkCounter,$val)"/> <!-- empty -->
                                <xsl:sequence select="d:sink(d:CopyToBaseCounter($myObj,$checkCounter))"/> <!-- empty -->
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:sequence select="d:sink(d:CopyToCurrCounter($myObj,$checknumId))"/> <!-- empty -->
                                <xsl:sequence select="d:sink(d:StartHeadingValueCtr($myObj,$checknumId,$val))"/> <!-- empty -->
                                <xsl:sequence select="d:sink(d:StartHeadingString($myObj,$checkilvl,$lStart,$checknumId,$val,'Vanish','No'))"/> <!-- empty -->
                                <xsl:sequence select="d:IncrementHeadingCounters($myObj,w:pPr/w:numPr/w:ilvl/@w:val,$checkCounter,$val)"/> <!-- empty -->
                                <xsl:sequence select="d:sink(d:CopyToBaseCounter($myObj,$checkCounter))"/> <!-- empty -->
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:when>
                </xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template name="HeadingsPart">
        <xsl:param name="checkilvl"/>
        <xsl:param name="checknumId"/>
        <xsl:param name="doc"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:variable name="val">
            <xsl:value-of select="$numberingXml//w:numbering/w:num[@w:numId=$checknumId]/w:abstractNumId/@w:val"/>
        </xsl:variable>

        <xsl:variable name="CheckNumId" select="d:CheckHeadingNumID($myObj,$checknumId)"/>
        <xsl:if test="$CheckNumId ='True'">
            <xsl:sequence select="d:sink(d:StartNewHeadingCounter($myObj,$checknumId,$val))"/> <!-- empty -->
        </xsl:if>

        <!--Checking numbering.xml for the type of List-->
        <xsl:variable    name="numFormat">
            <xsl:value-of select="$numberingXml//w:numbering/w:abstractNum[@w:abstractNumId=$val]/w:lvl[@w:ilvl=$checkilvl]/w:numFmt/@w:val"/>
        </xsl:variable>
        <xsl:variable    name="lvlText">
            <xsl:value-of select="$numberingXml//w:numbering/w:abstractNum[@w:abstractNumId=$val]/w:lvl[@w:ilvl=$checkilvl]/w:lvlText/@w:val"/>
        </xsl:variable>
        <xsl:variable    name="lStartOverride">
            <xsl:value-of select="$numberingXml//w:numbering/w:num[@w:numId=$checknumId]/w:lvlOverride[@w:ilvl=$checkilvl]/w:startOverride/@w:val"/>
        </xsl:variable>
        <xsl:variable    name="lStart">
            <xsl:value-of select="$numberingXml//w:numbering/w:abstractNum[@w:abstractNumId=$val]/w:lvl[@w:ilvl=$checkilvl]/w:start/@w:val"/>
        </xsl:variable>

        <xsl:sequence select="d:sink(d:AddCurrHeadId($myObj,$checknumId))"/> <!-- empty -->
        <xsl:variable name="addCurrLvl" select="d:AddCurrHeadLevel($myObj,$checkilvl,$doc,$val)"/>

        <xsl:choose>
            <xsl:when test="string-length(substring-before($addCurrLvl,'|'))=0">
                <xsl:choose>
                    <xsl:when test="not($lStartOverride='')">
                        <xsl:sequence select="d:sink(d:StartHeadingNewCtr($myObj,$checknumId,$val))"/> <!-- empty -->
                        <xsl:sequence select="d:sink(d:StartHeadingString($myObj,$checkilvl,$lStartOverride,$checknumId,$val,$doc,'Yes'))"/> <!-- empty -->
                        <xsl:sequence select="d:sink(d:CopyToCurrCounter($myObj,$checknumId))"/> <!-- empty -->
                    </xsl:when>
                    <xsl:when test="$lStartOverride='' and $numberingXml//w:numbering/w:num[@w:numId=$checknumId]/w:lvlOverride">
                        <xsl:sequence select="d:sink(d:StartHeadingNewCtr($myObj,$checknumId,$val))"/> <!-- empty -->
                        <xsl:sequence select="d:sink(d:StartHeadingString($myObj,$checkilvl,'0',$checknumId,$val,$doc,'Yes'))"/> <!-- empty -->
                        <xsl:sequence select="d:sink(d:CopyToCurrCounter($myObj,$checknumId))"/> <!-- empty -->
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:sequence select="d:sink(d:StartHeadingNewCtr($myObj,$checknumId,$val))"/> <!-- empty -->
                        <xsl:sequence select="d:sink(d:StartHeadingString($myObj,$checkilvl,$lStart,$checknumId,$val,$doc,'No'))"/> <!-- empty -->
                        <xsl:sequence select="d:sink(d:CopyToCurrCounter($myObj,$checknumId))"/> <!-- empty -->
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
        </xsl:choose>
        <xsl:choose>
            <xsl:when test="$doc='Document'">
                <xsl:sequence select="d:StoreHeadingPart($myObj,concat($numFormat,'|',$lvlText))"/> <!-- empty -->
            </xsl:when>
            <xsl:otherwise>
                <xsl:sequence select="d:StoreHeadingPart($myObj,concat($numFormat,'|',$lvlText,'!',$checknumId))"/> <!-- empty -->
            </xsl:otherwise>
        </xsl:choose>

    </xsl:template>

    <xsl:template name="CharacterStyles">
        <xsl:param name="characterStyle"/>
        <xsl:param name="txt"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:choose>
            <xsl:when test="$characterStyle='True'">
                <xsl:choose>
                    <xsl:when test="( w:rPr/w:u and(not(w:rPr/w:u[@w:val='none']))) and w:rPr/w:strike and w:rPr/w:color and w:rPr/w:sz and w:rPr/w:spacing">
                        <xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
                        <xsl:variable name="val" select="w:rPr/w:sz/@w:val"/>
                        <xsl:variable name="val_sz" select="($val)div 2"/>
                        <xsl:variable name="valspace" select="w:rPr/w:spacing/@w:val"/>
                        <xsl:variable name="val_spacing" select="($valspace*0.1)div 2"/>
                        <span class="{concat('text-decoration:Underline line-through;color:#',$val_color ,';letter-spacing:',$val_spacing ,';font-size:',$val_sz)}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="( w:rPr/w:u and(not(w:rPr/w:u[@w:val='none']))) and w:rPr/w:strike and w:rPr/w:color and w:rPr/w:sz">
                        <xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
                        <xsl:variable name="val" select="w:rPr/w:sz/@w:val"/>
                        <xsl:variable name="val_sz" select="($val)div 2"/>
                        <span class="{concat('text-decoration:Underline line-through;color:#',$val_color ,';font-size:',$val_sz)}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:strike and (w:rPr/w:u and(not(w:rPr/w:u[@w:val='none']))) and w:rPr/w:color">
                        <xsl:variable name="val" select="w:rPr/w:color/@w:val"/>
                        <span class="{concat('text-decoration:Underline line-through;color:#',$val)}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="( w:rPr/w:u and(not(w:rPr/w:u[@w:val='none']))) and w:rPr/w:color and w:rPr/w:sz">
                        <xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
                        <xsl:variable name="val" select="w:rPr/w:sz/@w:val"/>
                        <xsl:variable name="val_sz" select="($val)div 2"/>
                        <span class="{concat('text-decoration:Underline;color:#',$val_color,';font-size:',$val_sz)}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:strike and w:rPr/w:color and w:rPr/w:sz">
                        <xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
                        <xsl:variable name="val" select="w:rPr/w:sz/@w:val"/>
                        <xsl:variable name="val_sz" select="($val)div 2"/>
                        <span class="{concat('text-decoration:line-through;color:#',$val_color,';font-size:',$val_sz)}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:strike and ( w:rPr/w:u and(not(w:rPr/w:u[@w:val='none'])))">
                        <span class="text-decoration:Underline line-through">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="( w:rPr/w:u and(not(w:rPr/w:u[@w:val='none']))) and w:rPr/w:sz">
                        <xsl:variable name="val" select="w:rPr/w:sz/@w:val"/>
                        <xsl:variable name="val_sz" select="($val)div 2"/>
                        <span class="{concat('text-decoration:Underline;font-size:',$val_sz)}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:strike and w:rPr/w:sz">
                        <xsl:variable name="val" select="w:rPr/w:sz/@w:val"/>
                        <xsl:variable name="val_sz" select="($val)div 2"/>
                        <span class="{concat('text-decoration:line-through;font-size:',$val_sz)}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="( w:rPr/w:u and(not(w:rPr/w:u[@w:val='none']))) and w:rPr/w:color">
                        <xsl:variable name="val" select="w:rPr/w:color/@w:val"/>
                        <span class="{concat('text-decoration:Underline;color:#',$val)}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:strike and w:rPr/w:color">
                        <xsl:variable name="val" select="w:rPr/w:color/@w:val"/>
                        <span class="{concat('text-decoration:line-through;color:#',$val)}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:color and w:rPr/w:sz">
                        <xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
                        <xsl:variable name="val" select="w:rPr/w:sz/@w:val"/>
                        <xsl:variable name="val_sz" select="($val)div 2"/>
                        <span class="{concat('color:#',$val_color,';font-size:',$val_sz)}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:color and w:rPr/w:caps">
                        <xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
                        <span class="{concat('color:#',$val_color,';text-transform:uppercase')}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="( w:rPr/w:u and(not(w:rPr/w:u[@w:val='none']))) and w:rPr/w:caps">
                        <span class="text-decoration:Underline;'text-transform:uppercase'">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="( w:rPr/w:u and(not(w:rPr/w:u[@w:val='none']))) and w:rPr/w:smallCaps">
                        <span class="text-decoration:Underline;font-variant:small-caps">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:strike and w:rPr/w:caps">
                        <span class="text-decoration:line-through;text-transform:uppercase">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:strike and w:rPr/w:smallCaps">
                        <span class="text-decoration:line-through;font-variant:small-caps">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:color and w:rPr/w:smallCaps">
                        <xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
                        <span class="{concat('font-variant:small-caps,color:#',$val_color)}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>

                    <xsl:when test="( w:rPr/w:u and(not(w:rPr/w:u[@w:val='none'])))">
                        <span class="text-decoration:underline">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:strike">
                        <span class="text-decoration:line-through">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:smallCaps">
                        <span class="font-variant:small-caps">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:spacing">
                        <xsl:variable name="val" select="w:rPr/w:spacing/@w:val"/>
                        <xsl:variable name="val_spacing" select="($val*0.1)div 2"/>
                        <span class="{concat('letter-spacing:',$val_spacing,'pt')}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:color">
                        <xsl:variable name="val" select="w:rPr/w:color/@w:val"/>
                        <span class="{concat('color:#',$val)}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:caps">
                        <span class="text-transform:uppercase">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:rPr/w:sz">
                        <xsl:variable name="val" select="w:rPr/w:sz/@w:val"/>
                        <xsl:variable name="val_sz" select="($val)div 2"/>
                        <span class="{concat('font-size:',$val_sz)}">
                            <xsl:if test="not($txt='')">
                                <xsl:value-of select="$txt"/>
                            </xsl:if>
                            <xsl:value-of select="w:t"/>
                        </span>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:if test="not($txt='')">
                            <xsl:value-of select="$txt"/>
                        </xsl:if>
                        <xsl:value-of select="w:t"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
                <xsl:if test="not($txt='')">
                    <xsl:value-of select="$txt"/>
                </xsl:if>
                <xsl:value-of select="w:t"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <!--Template to implement Custom and inbuilt paragraph styles.-->
    <xsl:template name="ParagraphStyle">
        <xsl:param name="prmTrack"/>
        <xsl:param name="VERSION"/>
        <xsl:param name="flag"/>
        <xsl:param name="flagNote"/>
        <xsl:param name="checkid"/>
        <xsl:param name="custom"/>
        <xsl:param name="sOperators"/>
        <xsl:param name="sMinuses"/>
        <xsl:param name="sNumbers"/>
        <xsl:param name="sZeros"/>
        <xsl:param name="txt"/>
        <xsl:param name="masterparastyle"/>
        <xsl:param name="imgOptionPara"/>
        <xsl:param name="dpiPara"/>
        <xsl:param name="characterparaStyle"/>

        <xsl:variable name="aquote">"</xsl:variable>
        <xsl:message terminate="no">progress:ParagraphStyle</xsl:message>
        <xsl:variable name="checkImageposition" select="d:GetCaptionsProdnotes($myObj)"/>
        <xsl:choose>
            <!--Checking for Title/Subtitle paragraph style-->
            <xsl:when test="(w:pPr/w:pStyle/@w:val='Title') or (w:pPr/w:pStyle/@w:val='Subtitle')">
                <xsl:variable name="lang">
                    <xsl:call-template name="Languages">
                        <xsl:with-param name="Attribute" select="'1'"/>
                    </xsl:call-template>
                </xsl:variable>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','doctitle ','xml:lang=',$aquote,$lang,$aquote,'&gt;')"/>
                <xsl:call-template name="ParaHandler">
                    <xsl:with-param name="flag" select="'0'"/>
                    <xsl:with-param name="VERSION" select="$VERSION"/>
                    <xsl:with-param name="custom" select="$custom"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="charparahandlerStyle" select="$characterparaStyle"/>
                </xsl:call-template>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/doctitle','&gt;')"/>
            </xsl:when>
            <!--Checking for AuthorDAISY custom paragraph style-->
            <xsl:when test="(w:pPr/w:pStyle/@w:val='AuthorDAISY')">
                <xsl:variable name="lang">
                    <xsl:call-template name="Languages">
                        <xsl:with-param name="Attribute" select="'1'"/>
                    </xsl:call-template>
                </xsl:variable>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','author ','xml:lang=',$aquote,$lang,$aquote,'&gt;')"/>
                <xsl:if test="$flagNote='Note'">
                    <xsl:if test="d:NoteFlag($myObj)=1">
                        <p>
                            <xsl:value-of select="$checkid - 1"/>
                        </p>
                    </xsl:if>
                </xsl:if>
                <xsl:call-template name="Paracharacterstyle">
                    <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="flag" select="0"/>
                </xsl:call-template>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/author','&gt;')"/>
            </xsl:when>
            <!--Checking for CovertitleDAISY custom paragraph style-->
            <xsl:when test="(w:pPr/w:pStyle/@w:val='CovertitleDAISY')">
                <xsl:variable name="lang">
                    <xsl:call-template name="Languages">
                        <xsl:with-param name="Attribute" select="'1'"/>
                    </xsl:call-template>
                </xsl:variable>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','covertitle ','xml:lang=',$aquote,$lang,$aquote,'&gt;')"/>
                <xsl:call-template name="Paracharacterstyle">
                    <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="flag" select="0"/>
                </xsl:call-template>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/covertitle','&gt;')"/>
            </xsl:when>
            <!--Checking for BylineDAISY custom paragraph style-->
            <xsl:when test="(w:pPr/w:pStyle/@w:val='BylineDAISY')">
                <xsl:if test="$flagNote='Note'">
                    <xsl:if test="d:NoteFlag($myObj)=1">
                        <p>
                            <xsl:value-of select="$checkid - 1"/>
                        </p>
                    </xsl:if>
                </xsl:if>
                <xsl:variable name="lang">
                    <xsl:call-template name="Languages">
                        <xsl:with-param name="Attribute" select="'1'"/>
                    </xsl:call-template>
                </xsl:variable>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','byline ','xml:lang=',$aquote,$lang,$aquote,'&gt;')"/>
                <xsl:call-template name="Paracharacterstyle">
                    <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="flag" select="0"/>
                </xsl:call-template>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/byline','&gt;')"/>
            </xsl:when>
            <!--Checking for DatelineDAISY custom paragraph style-->
            <xsl:when test="(w:pPr/w:pStyle/@w:val='DatelineDAISY')">
                <xsl:if test="$flagNote='Note'">
                    <xsl:if test="d:NoteFlag($myObj)=1">
                        <p>
                            <xsl:value-of select="$checkid - 1"/>
                        </p>
                    </xsl:if>
                </xsl:if>
                <xsl:variable name="lang">
                    <xsl:call-template name="Languages">
                        <xsl:with-param name="Attribute" select="'1'"/>
                    </xsl:call-template>
                </xsl:variable>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','dateline ','xml:lang=',$aquote,$lang,$aquote,'&gt;')"/>
                <xsl:call-template name="Paracharacterstyle">
                    <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="flag" select="0"/>
                </xsl:call-template>
                <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/dateline','&gt;')"/>
            </xsl:when>
            <!--Checking for Prodnote-OptionalDAISY custom paragraph style-->
            <xsl:when test="(w:pPr/w:pStyle/@w:val='Prodnote-OptionalDAISY')and (not((preceding-sibling::node()[$checkImageposition]/w:r/w:drawing) or (preceding-sibling::node()[$checkImageposition]/w:r/w:pict) or (preceding-sibling::node()[$checkImageposition]/w:r/w:object)))">
                <xsl:if test="count(preceding-sibling::node()[1]/w:pPr/w:pStyle[@w:val='Prodnote-OptionalDAISY'])=0">
                    <xsl:variable name="lang">
                        <xsl:call-template name="Languages">
                            <xsl:with-param name="Attribute" select="'1'"/>
                        </xsl:call-template>
                    </xsl:variable>
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','prodnote ','render=',$aquote,'optional',$aquote,' xml:lang=',$aquote,$lang,$aquote,'&gt;')"/>
                </xsl:if>

                <xsl:call-template name="Paracharacterstyle">
                    <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="flag" select="1"/>
                </xsl:call-template>

                <xsl:if test="count(following-sibling::node()[1]/w:pPr/w:pStyle[contains(@w:val,'Prodnote-OptionalDAISY')])=0">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/prodnote ','&gt;')"/>
                </xsl:if>
            </xsl:when>
            <!--Checking for Prodnote-RequiredDAISY custom paragraph style-->
            <xsl:when test="(w:pPr/w:pStyle/@w:val='Prodnote-RequiredDAISY')and (not((preceding-sibling::node()[$checkImageposition]/w:r/w:drawing) or (preceding-sibling::node()[$checkImageposition]/w:r/w:pict) or (preceding-sibling::node()[$checkImageposition]/w:r/w:object)))">
                <xsl:if test="count(preceding-sibling::node()[1]/w:pPr/w:pStyle[@w:val='Prodnote-RequiredDAISY'])=0">
                    <xsl:variable name="lang">
                        <xsl:call-template name="Languages">
                            <xsl:with-param name="Attribute" select="'1'"/>
                        </xsl:call-template>
                    </xsl:variable>
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','prodnote ','render=',$aquote,'required',$aquote,' xml:lang=',$aquote,$lang,$aquote,'&gt;')"/>
                </xsl:if>

                <xsl:call-template name="Paracharacterstyle">
                    <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="flag" select="1"/>
                </xsl:call-template>

                <xsl:if test="count(following-sibling::node()[1]/w:pPr/w:pStyle[contains(@w:val,'Prodnote-RequiredDAISY')])=0">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/prodnote','&gt;')"/>
                </xsl:if>
            </xsl:when>
            <!--Checking for PoemDAISY/Poem-TitleDAISY/Poem-HeadingDAISY/Poem-AuthorDAISY/Poem-BylineDAISY custom paragraph styles-->
            <xsl:when test="w:pPr/w:pStyle[substring(@w:val,1,4)='Poem']">
                <xsl:if test="$flagNote='Note'">
                    <xsl:if test="d:NoteFlag($myObj)=1">
                        <p>
                            <xsl:value-of select="$checkid - 1"/>
                        </p>
                    </xsl:if>
                </xsl:if>
                <xsl:if test="count(preceding-sibling::node()[1]/w:pPr/w:pStyle[substring(@w:val,1,4)='Poem'])=0">
                    <xsl:variable name="lang">
                        <xsl:call-template name="Languages">
                            <xsl:with-param name="Attribute" select="'1'"/>
                        </xsl:call-template>
                    </xsl:variable>
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','poem ','xml:lang=',$aquote,$lang,$aquote,'&gt;')"/>
                </xsl:if>
                <xsl:if test="w:pPr/w:pStyle/@w:val='Poem-TitleDAISY'">
                    <title>
                        <xsl:call-template name="Paracharacterstyle">
                            <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                            <xsl:with-param name="txt" select="$txt"/>
                            <xsl:with-param name="flag" select="0"/>
                        </xsl:call-template>
                    </title>
                </xsl:if>
                <xsl:if test="w:pPr/w:pStyle/@w:val='Poem-HeadingDAISY'">
                    <hd>
                        <xsl:call-template name="Paracharacterstyle">
                            <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                            <xsl:with-param name="txt" select="$txt"/>
                            <xsl:with-param name="flag" select="0"/>
                        </xsl:call-template>
                    </hd>
                </xsl:if>
                <xsl:if test="(w:pPr/w:pStyle/@w:val='PoemDAISY')">
                    <xsl:if test="count(preceding-sibling::node()[1]/w:pPr/w:pStyle[@w:val='PoemDAISY'])=0">
                        <xsl:variable name="lang">
                            <xsl:call-template name="Languages">
                                <xsl:with-param name="Attribute" select="'1'"/>
                            </xsl:call-template>
                        </xsl:variable>
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','linegroup ','xml:lang=',$aquote,$lang,$aquote,'&gt;')"/>
                    </xsl:if>
                    <line>
                        <xsl:call-template name="Paracharacterstyle">
                            <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                            <xsl:with-param name="txt" select="$txt"/>
                            <xsl:with-param name="flag" select="0"/>
                        </xsl:call-template>
                    </line>
                    <xsl:if test="count(following-sibling::node()[1]/w:pPr/w:pStyle[@w:val='PoemDAISY'])=0">
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/linegroup','&gt;')"/>
                    </xsl:if>
                </xsl:if>
                <xsl:if test="(w:pPr/w:pStyle/@w:val='Poem-AuthorDAISY')">
                    <author>
                        <xsl:call-template name="Paracharacterstyle">
                            <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                            <xsl:with-param name="txt" select="$txt"/>
                            <xsl:with-param name="flag" select="0"/>
                        </xsl:call-template>
                    </author>
                </xsl:if>
                <xsl:if test="(w:pPr/w:pStyle/@w:val='Poem-BylineDAISY')">
                    <byline>
                        <xsl:call-template name="Paracharacterstyle">
                            <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                            <xsl:with-param name="txt" select="$txt"/>
                            <xsl:with-param name="flag" select="0"/>
                        </xsl:call-template>
                    </byline>
                </xsl:if>
                <xsl:if test="count(following-sibling::node()[1]/w:pPr/w:pStyle[substring(@w:val,1,4)='Poem'])=0">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/poem','&gt;')"/>
                </xsl:if>
            </xsl:when>
            <!--Checking for EpigraphDAISY/Epigraph-AuthorDAISY custom paragraph styles-->
            <xsl:when test="(w:pPr/w:pStyle[substring(@w:val,1,8)='Epigraph'])">
                <xsl:if test="count(preceding-sibling::node()[1]/w:pPr/w:pStyle[substring(@w:val,1,8)='Epigraph'])=0">
                    <xsl:variable name="lang">
                        <xsl:call-template name="Languages">
                            <xsl:with-param name="Attribute" select="'1'"/>
                        </xsl:call-template>
                    </xsl:variable>
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','epigraph ','xml:lang=',$aquote,$lang,$aquote,'&gt;')"/>
                </xsl:if>
                <xsl:call-template name="Paracharacterstyle">
                    <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="flag" select="1"/>
                </xsl:call-template>
                <xsl:if test="count(following-sibling::node()[1]/w:pPr/w:pStyle[substring(@w:val,1,8)='Epigraph'])=0">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/epigraph','&gt;')"/>
                </xsl:if>
            </xsl:when>
            <!--Checking for AddressDAISY custom paragraph style-->
            <xsl:when test="w:pPr/w:pStyle[@w:val='AddressDAISY']">
                <xsl:if test="$flagNote='Note'">
                    <xsl:if test="d:NoteFlag($myObj)=1">
                        <p>
                            <xsl:value-of select="$checkid - 1"/>
                        </p>
                    </xsl:if>
                </xsl:if>
                <xsl:choose>
                    <!--Checking for occurence of Lists in AddressDAISY custom paragraph style-->
                    <xsl:when test="(w:pPr/w:numPr/w:ilvl) and (w:pPr/w:numPr/w:numId)">
                        <xsl:variable name="lang">
                            <xsl:call-template name="Languages">
                                <xsl:with-param name="Attribute" select="'1'"/>
                            </xsl:call-template>
                        </xsl:variable>
                        <!--Opening Address tag-->
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','address ','xml:lang=',$aquote,$lang,$aquote,'&gt;')"/>
                    </xsl:when>
                    <xsl:when test="count(preceding-sibling::node()[1]/w:pPr/w:pStyle[@w:val='AddressDAISY'])=0">
                        <xsl:variable name="lang">
                            <xsl:call-template name="Languages">
                                <xsl:with-param name="Attribute" select="'1'"/>
                            </xsl:call-template>
                        </xsl:variable>
                        <!--Opening Address tag-->
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','address ','xml:lang=',$aquote,$lang,$aquote,'&gt;')"/>
                    </xsl:when>
                </xsl:choose>
                <line>
                    <xsl:call-template name="Paracharacterstyle">
                        <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                        <xsl:with-param name="txt" select="$txt"/>
                        <xsl:with-param name="flag" select="0"/>
                    </xsl:call-template>
                </line>
                <xsl:choose>
                    <!--Checking for Address in a list-->
                    <xsl:when test="(w:pPr/w:numPr/w:ilvl) and (w:pPr/w:numPr/w:numId)">
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/address','&gt;')"/>
                    </xsl:when>
                    <!--Checking for address style in the next sibling and closing the tag-->
                    <xsl:when test="count(following-sibling::node()[1]/w:pPr/w:pStyle[@w:val='AddressDAISY'])=0">
                        <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/address','&gt;')"/>
                    </xsl:when>
                </xsl:choose>
            </xsl:when>
            <!--Checking for DivDAISY custom paragraph style-->
            <xsl:when test="w:pPr/w:pStyle[substring(@w:val,1,3)='Div']">
                <xsl:if test="count(preceding-sibling::node()[1]/w:pPr/w:pStyle[substring(@w:val,1,3)='Div'])=0">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','div','&gt;')"/>
                </xsl:if>

                <xsl:call-template name="Paracharacterstyle">
                    <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="flag" select="1"/>
                </xsl:call-template>

                <xsl:if test="count(following-sibling::node()[1]/w:pPr/w:pStyle[substring(@w:val,1,3)='Div'])=0">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/div','&gt;')"/>
                </xsl:if>
            </xsl:when>

            <!--Checking for occurence of Lists in paragraph-->
            <xsl:when test="(w:pPr/w:numPr/w:ilvl) and (w:pPr/w:numPr/w:numId)">
                <xsl:call-template name="ParaHandler">
                    <xsl:with-param name="flag" select="'3'"/>
                    <xsl:with-param name="VERSION" select="$VERSION"/>
                    <xsl:with-param name="flagNote" select="$flagNote"/>
                    <xsl:with-param name="custom" select="$custom"/>
                    <xsl:with-param name="sOperators" select="$sOperators"/>
                    <xsl:with-param name="sMinuses" select="$sMinuses"/>
                    <xsl:with-param name="sNumbers" select="$sNumbers"/>
                    <xsl:with-param name="sZeros" select="$sZeros"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="mastersubpara" select="$masterparastyle"/>
                    <xsl:with-param name="charparahandlerStyle" select="$characterparaStyle"/>
                </xsl:call-template>
            </xsl:when>

            <!--Checking for DefinitionTermDAISY custom character style and DefinitionDataDAISY    custom paragraph style-->
            <xsl:when test="(w:r/w:rPr/w:rStyle/@w:val='DefinitionTermDAISY') or (w:pPr/w:pStyle/@w:val='DefinitionDataDAISY')">
                <xsl:if test="(count(preceding-sibling::node()[1]/w:pPr/w:pStyle[@w:val='DefinitionDataDAISY'])=0)
                                                                                        and (count(preceding-sibling::node()[1]/w:r/w:rPr/w:rStyle[@w:val='DefinitionTermDAISY'])=0)">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','dl','&gt;')"/>
                </xsl:if>
                <!--Checking for DefinitionTermDAISY custom character style-->
                <xsl:if test="w:r/w:rPr/w:rStyle/@w:val='DefinitionTermDAISY'">
                    <dt>
                        <!--Checking if image is bidirectionally oriented-->
                        <xsl:if test="(w:pPr/w:bidi) or (w:r/w:rPr/w:rtl)">
                            <!--Variable holds the value which indicates that the image is bidirectionally oriented-->
                            <xsl:variable name="definitionlistBd">
                                <xsl:call-template name="BdoRtlLanguages"/>
                            </xsl:variable>
                            <xsl:variable name="quote">"</xsl:variable>
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$definitionlistBd,$quote,'&gt;')"/>
                        </xsl:if>
                        <xsl:value-of select="w:r/w:t"/>
                        <xsl:if test="(w:pPr/w:bidi) or (w:r/w:rPr/w:rtl)">
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
                        </xsl:if>
                    </dt>
                </xsl:if>
                <!--Checking for DefinitionData-DAISY custom character style-->
                <xsl:if test="(w:pPr/w:pStyle/@w:val='DefinitionDataDAISY')">
                    <dd>
                        <xsl:call-template name="Paracharacterstyle">
                            <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                            <xsl:with-param name="txt" select="$txt"/>
                            <xsl:with-param name="flag" select="0"/>
                        </xsl:call-template>
                    </dd>
                </xsl:if>
                <xsl:if test="(count(following-sibling::node()[1]/w:pPr/w:pStyle[@w:val='DefinitionDataDAISY'])=0)
                                                                                        and (count(following-sibling::node()[1]/w:r/w:rPr/w:rStyle[@w:val='DefinitionTermDAISY'])=0)">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/dl','&gt;')"/>
                </xsl:if>
            </xsl:when>

            <xsl:when test="$characterparaStyle='True'">
                <xsl:choose>
                    <xsl:when test="w:pPr/w:ind[@w:left] and w:pPr/w:ind[@w:right]">
                        <p>
                            <xsl:variable name="val" select="w:pPr/w:ind/@w:left"/>
                            <xsl:variable name="val_left" select="($val div 1440)"/>
                            <xsl:variable name="valright" select="w:pPr/w:ind/@w:right"/>
                            <xsl:variable name="val_right" select="($valright div 1440)"/>
                            <span class="{concat('text-indent:', 'right=',$val_right,'in',';left=',$val_left,'in')}">
                                <xsl:call-template name="ParaHandler">
                                    <xsl:with-param name="flag" select="'0'"/>
                                    <xsl:with-param name="VERSION" select="$VERSION"/>
                                    <xsl:with-param name="custom" select="$custom"/>
                                    <xsl:with-param name="txt" select="$txt"/>
                                    <xsl:with-param name="charparahandlerStyle" select="$characterparaStyle"/>
                                </xsl:call-template>
                            </span>
                        </p>
                    </xsl:when>
                    <xsl:when test="w:pPr/w:ind[@w:left] and    w:pPr/w:jc">
                        <p>
                            <xsl:variable name="val" select="w:pPr/w:ind/@w:left"/>
                            <xsl:variable name="val_left" select="($val div 1440)"/>
                            <xsl:variable name="val1" select="w:pPr/w:jc/@w:val"/>
                            <span class="{concat('text-indent:',';left=',$val_left,'in',';text-align:',$val1)}">
                                <xsl:call-template name="ParaHandler">
                                    <xsl:with-param name="flag" select="'0'"/>
                                    <xsl:with-param name="VERSION" select="$VERSION"/>
                                    <xsl:with-param name="custom" select="$custom"/>
                                    <xsl:with-param name="txt" select="$txt"/>
                                    <xsl:with-param name="charparahandlerStyle" select="$characterparaStyle"/>
                                </xsl:call-template>
                            </span>
                        </p>
                    </xsl:when>
                    <xsl:when test="w:pPr/w:ind[@w:left]">
                        <p>
                            <xsl:variable name="val" select="w:pPr/w:ind/@w:left"/>
                            <xsl:variable name="val_left" select="($val div 1440)"/>
                            <span class="{concat('text-indent:',$val_left,'in')}">
                                <xsl:call-template name="ParaHandler">
                                    <xsl:with-param name="flag" select="'0'"/>
                                    <xsl:with-param name="VERSION" select="$VERSION"/>
                                    <xsl:with-param name="custom" select="$custom"/>
                                    <xsl:with-param name="charparahandlerStyle" select="$characterparaStyle"/>
                                </xsl:call-template>
                            </span>
                        </p>
                    </xsl:when>
                    <xsl:when test="w:pPr/w:ind[@w:right]">
                        <p>
                            <xsl:variable name="val" select="w:pPr/w:ind/@w:right"/>
                            <xsl:variable name="val_right" select="($val div 1440)"/>
                            <span class="{concat('text-indent:',$val_right,'in')}">
                                <xsl:call-template name="ParaHandler">
                                    <xsl:with-param name="flag" select="'0'"/>
                                    <xsl:with-param name="VERSION" select="$VERSION"/>
                                    <xsl:with-param name="custom" select="$custom"/>
                                    <xsl:with-param name="txt" select="$txt"/>
                                    <xsl:with-param name="charparahandlerStyle" select="$characterparaStyle"/>
                                </xsl:call-template>
                            </span>
                        </p>
                    </xsl:when>
                    <xsl:when test="w:pPr/w:jc">
                        <p>
                            <xsl:variable name="val" select="w:pPr/w:jc/@w:val"/>
                            <span class="{concat('text-align:',$val)}">
                                <xsl:call-template name="ParaHandler">
                                    <xsl:with-param name="flag" select="'0'"/>
                                    <xsl:with-param name="VERSION" select="$VERSION"/>
                                    <xsl:with-param name="custom" select="$custom"/>
                                    <xsl:with-param name="txt" select="$txt"/>
                                    <xsl:with-param name="charparahandlerStyle" select="$characterparaStyle"/>
                                </xsl:call-template>
                            </span>
                        </p>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:call-template name="ParaHandler">
                            <xsl:with-param name="flag" select="'1'"/>
                            <xsl:with-param name="prmTrack" select="$prmTrack"/>
                            <xsl:with-param name="VERSION" select="$VERSION"/>
                            <xsl:with-param name="flagNote" select="$flagNote"/>
                            <xsl:with-param name="checkid" select="$checkid"/>
                            <xsl:with-param name="custom" select="$custom"/>
                            <xsl:with-param name="mastersubpara" select="$masterparastyle"/>
                            <xsl:with-param name="imgOptionPara" select="$imgOptionPara"/>
                            <xsl:with-param name="dpiPara" select="$dpiPara"/>
                            <xsl:with-param name="txt" select="$txt"/>
                            <xsl:with-param name="charparahandlerStyle" select="$characterparaStyle"/>
                        </xsl:call-template>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>

            <!--Checking for occurence of Lists in paragraph-->
            <xsl:when test="(w:pPr/w:numPr/w:ilvl) and (w:pPr/w:numPr/w:numId)">
                <xsl:call-template name="ParaHandler">
                    <xsl:with-param name="flag" select="'3'"/>
                    <xsl:with-param name="VERSION" select="$VERSION"/>
                    <xsl:with-param name="flagNote" select="$flagNote"/>
                    <xsl:with-param name="custom" select="$custom"/>
                    <xsl:with-param name="sOperators" select="$sOperators"/>
                    <xsl:with-param name="sMinuses" select="$sMinuses"/>
                    <xsl:with-param name="sNumbers" select="$sNumbers"/>
                    <xsl:with-param name="sZeros" select="$sZeros"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="mastersubpara" select="$masterparastyle"/>
                    <xsl:with-param name="charparahandlerStyle" select="$characterparaStyle"/>
                </xsl:call-template>
            </xsl:when>
            <!--Checking for DefinitionTermDAISY custom character style and DefinitionDataDAISY    custom paragraph style-->
            <xsl:when test="(w:r/w:rPr/w:rStyle/@w:val='DefinitionTermDAISY') or (w:pPr/w:pStyle/@w:val='DefinitionDataDAISY')">
                <xsl:if test="(count(preceding-sibling::node()[1]/w:pPr/w:pStyle[@w:val='DefinitionDataDAISY'])=0)
                                                                                        and (count(preceding-sibling::node()[1]/w:r/w:rPr/w:rStyle[@w:val='DefinitionTermDAISY'])=0)">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','dl','&gt;')"/>
                </xsl:if>
                <!--Checking for DefinitionTermDAISY custom character style-->
                <xsl:if test="w:r/w:rPr/w:rStyle/@w:val='DefinitionTermDAISY'">
                    <dt>
                        <!--Checking if image is bidirectionally oriented-->
                        <xsl:if test="(w:pPr/w:bidi) or (w:r/w:rPr/w:rtl)">
                            <!--Variable holds the value which indicates that the image is bidirectionally oriented-->
                            <xsl:variable name="definitionlistBd">
                                <xsl:call-template name="BdoRtlLanguages"/>
                            </xsl:variable>
                            <xsl:variable name="quote">"</xsl:variable>
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$definitionlistBd,$quote,'&gt;')"/>
                        </xsl:if>
                        <xsl:value-of select="w:r/w:t"/>
                        <xsl:if test="(w:pPr/w:bidi) or (w:r/w:rPr/w:rtl)">
                            <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
                        </xsl:if>
                    </dt>
                </xsl:if>
                <!--Checking for DefinitionData-DAISY custom character style-->
                <xsl:if test="(w:pPr/w:pStyle/@w:val='DefinitionDataDAISY')">
                    <dd>
                        <xsl:call-template name="Paracharacterstyle">
                            <xsl:with-param name="characterStyle" select="$characterparaStyle"/>
                            <xsl:with-param name="txt" select="$txt"/>
                            <xsl:with-param name="flag" select="0"/>
                        </xsl:call-template>
                    </dd>
                </xsl:if>
                <xsl:if test="(count(following-sibling::node()[1]/w:pPr/w:pStyle[@w:val='DefinitionDataDAISY'])=0)
                                                                                        and (count(following-sibling::node()[1]/w:r/w:rPr/w:rStyle[@w:val='DefinitionTermDAISY'])=0)">
                    <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/dl','&gt;')"/>
                </xsl:if>
            </xsl:when>
            <!--Checking if table occurs in the document and implementing all the styles applied on it-->
            <xsl:when test="(name(..)='w:tc')
                                                                                        and (not((w:pPr/w:pStyle/@w:val='Prodnote-OptionalDAISY')
                                                                                                                        or (w:pPr/w:pStyle/@w:val='Prodnote-RequiredDAISY')
                                                                                                                        or (w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')))">
                <xsl:call-template name="ParaHandler">
                    <xsl:with-param name="flag" select="'2'"/>
                    <xsl:with-param name="VERSION" select="$VERSION"/>
                    <xsl:with-param name="mastersubpara" select="$masterparastyle"/>
                    <xsl:with-param name="custom" select="$custom"/>
                    <xsl:with-param name="charparahandlerStyle" select="$characterparaStyle"/>
                </xsl:call-template>
            </xsl:when>
            <!--Checking if no style exist and then calling the template named ParaHandler-->
            <xsl:when test="not((((w:pPr/w:pStyle/@w:val='Table-CaptionDAISY')
                                                                                                                or (w:pPr/w:pStyle/@w:val='Caption')
                                                                                                                or (./node()[name()='w:fldSimple']))
                                                                                                        and ((preceding-sibling::node()[1][name()='w:tbl'])
                                                                                                                or (following-sibling::node()[1][name()='w:tbl'])))
                                                                                                or (w:pPr/w:pStyle[substring(@w:val,1,3)='TOC'])
                                                                                                or (preceding-sibling::node()[$checkImageposition]/w:r/w:drawing)
                                                                                                or (preceding-sibling::node()[$checkImageposition]/w:r/w:pict)
                                                                                                or ((w:pPr/w:pStyle[@w:val='Image-CaptionDAISY'])
                                                                                                        and ((following-sibling::node()[1]/w:r/w:drawing)
                                                                                                                or (following-sibling::node()[1]/w:r/w:pict))))">
                <xsl:call-template name="ParaHandler">
                    <xsl:with-param name="flag" select="'1'"/>
                    <xsl:with-param name="prmTrack" select="$prmTrack"/>
                    <xsl:with-param name="VERSION" select="$VERSION"/>
                    <xsl:with-param name="flagNote" select="$flagNote"/>
                    <xsl:with-param name="checkid" select="$checkid"/>
                    <xsl:with-param name="custom" select="$custom"/>
                    <xsl:with-param name="mastersubpara" select="$masterparastyle"/>
                    <xsl:with-param name="imgOptionPara" select="$imgOptionPara"/>
                    <xsl:with-param name="dpiPara" select="$dpiPara"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="charparahandlerStyle" select="$characterparaStyle"/>
                </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
                <!--Other elements are considered as fidelity loss-->
                <!--Capturing fidility loss elements-->
                <xsl:if test="not((name()='w:pPr')
                                                                                                or (name()='w:p')
                                                                                                or (name()='w:r')
                                                                                                or (name()='w:fldSimple')
                                                                                                or (name()='w:fldChar')
                                                                                                or (name()='w:proofErr')
                                                                                                or (name()='w:lastRenderedPageBreak')
                                                                                                or (name()='w:br')
                                                                                                or (name()='w:tab'))">
                    <xsl:message terminate="no">
                        <xsl:value-of select="concat('translation.oox2Daisy.UncoveredElement|', name())"/>
                    </xsl:message>
                </xsl:if>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template name="Paracharacterstyle">
        <xsl:param name="VERSION"/>
        <xsl:param name="custom"/>
        <xsl:param name="characterStyle"/>
        <xsl:param name="flag"/>
        <xsl:param name="txt"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
        <xsl:choose>
            <xsl:when test="$characterStyle='True'">
                <xsl:choose>
                    <xsl:when test="w:pPr/w:ind[@w:left] and w:pPr/w:ind[@w:right]">
                        <xsl:variable name="val" select="w:pPr/w:ind/@w:left"/>
                        <xsl:variable name="val_left" select="($val div 1440)"/>
                        <xsl:variable name="valright" select="w:pPr/w:ind/@w:right"/>
                        <xsl:variable name="val_right" select="($valright div 1440)"/>
                        <span class="{concat('text-indent:', 'right=',$val_right,'in',';left=',$val_left,'in')}">
                            <xsl:call-template name="ParaHandler">
                                <xsl:with-param name="flag" select="'0'"/>
                                <xsl:with-param name="VERSION" select="$VERSION"/>
                                <xsl:with-param name="custom" select="$custom"/>
                                <xsl:with-param name="txt" select="$txt"/>
                                <xsl:with-param name="charparahandlerStyle" select="$characterStyle"/>
                            </xsl:call-template>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:pPr/w:ind[@w:left] and    w:pPr/w:jc">
                        <xsl:variable name="val" select="w:pPr/w:ind/@w:left"/>
                        <xsl:variable name="val_left" select="($val div 1440)"/>
                        <xsl:variable name="val1" select="w:pPr/w:jc/@w:val"/>
                        <span class="{concat('text-indent:',';left=',$val_left,'in',';text-align:',$val1)}">
                            <xsl:call-template name="ParaHandler">
                                <xsl:with-param name="flag" select="'0'"/>
                                <xsl:with-param name="VERSION" select="$VERSION"/>
                                <xsl:with-param name="custom" select="$custom"/>
                                <xsl:with-param name="txt" select="$txt"/>
                                <xsl:with-param name="charparahandlerStyle" select="$characterStyle"/>
                            </xsl:call-template>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:pPr/w:ind[@w:left]">
                        <xsl:variable name="val" select="w:pPr/w:ind/@w:left"/>
                        <xsl:variable name="val_left" select="($val div 1440)"/>
                        <span class="{concat('text-indent:',$val_left,'in')}">
                            <xsl:call-template name="ParaHandler">
                                <xsl:with-param name="flag" select="'0'"/>
                                <xsl:with-param name="VERSION" select="$VERSION"/>
                                <xsl:with-param name="custom" select="$custom"/>
                                <xsl:with-param name="txt" select="$txt"/>
                                <xsl:with-param name="charparahandlerStyle" select="$characterStyle"/>
                            </xsl:call-template>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:pPr/w:ind[@w:right]">
                        <xsl:variable name="val" select="w:pPr/w:ind/@w:right"/>
                        <xsl:variable name="val_right" select="($val div 1440)"/>
                        <span class="{concat('text-indent:',$val_right,'in')}">
                            <xsl:call-template name="ParaHandler">
                                <xsl:with-param name="flag" select="'0'"/>
                                <xsl:with-param name="VERSION" select="$VERSION"/>
                                <xsl:with-param name="custom" select="$custom"/>
                                <xsl:with-param name="txt" select="$txt"/>
                                <xsl:with-param name="charparahandlerStyle" select="$characterStyle"/>
                            </xsl:call-template>
                        </span>
                    </xsl:when>
                    <xsl:when test="w:pPr/w:jc">
                        <xsl:variable name="val" select="w:pPr/w:jc/@w:val"/>
                        <span class="{concat('text-align:',$val)}">
                            <xsl:call-template name="ParaHandler">
                                <xsl:with-param name="flag" select="'0'"/>
                                <xsl:with-param name="VERSION" select="$VERSION"/>
                                <xsl:with-param name="custom" select="$custom"/>
                                <xsl:with-param name="txt" select="$txt"/>
                                <xsl:with-param name="charparahandlerStyle" select="$characterStyle"/>
                            </xsl:call-template>
                        </span>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:call-template name="ParaHandler">
                            <xsl:with-param name="flag" select="$flag"/>
                            <xsl:with-param name="VERSION" select="$VERSION"/>
                            <xsl:with-param name="custom" select="$custom"/>
                            <xsl:with-param name="txt" select="$txt"/>
                            <xsl:with-param name="charparahandlerStyle" select="$characterStyle"/>
                        </xsl:call-template>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
                <xsl:call-template name="ParaHandler">
                    <xsl:with-param name="flag" select="$flag"/>
                    <xsl:with-param name="VERSION" select="$VERSION"/>
                    <xsl:with-param name="custom" select="$custom"/>
                    <xsl:with-param name="txt" select="$txt"/>
                    <xsl:with-param name="charparahandlerStyle" select="$characterStyle"/>
                </xsl:call-template>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template name="CheckPageStyles">
        <xsl:for-each select="$documentXml//w:document/w:body/node()">
            <xsl:message terminate="no">progress:checkpagestyles</xsl:message>
            <xsl:if test="name()='w:p'">
                <xsl:for-each select="w:pPr/w:pStyle[substring(@w:val,1,11)='Frontmatter']">
                    <xsl:if test="d:PushPageStyle($myObj,@w:val)"/>
                </xsl:for-each>
                <xsl:for-each select="w:pPr/w:pStyle[substring(@w:val,1,10)='Bodymatter']">
                    <xsl:if test="d:PushPageStyle($myObj,@w:val)"/>
                </xsl:for-each>
                <xsl:for-each select="w:pPr/w:pStyle[substring(@w:val,1,10)='Rearmatter']">
                    <xsl:if test="d:PushPageStyle($myObj,@w:val)"/>
                </xsl:for-each>
                <xsl:for-each select="w:r/w:rPr/w:rStyle[substring(@w:val,1,11)='Frontmatter']">
                    <xsl:if test="d:PushPageStyle($myObj,@w:val)"/>
                </xsl:for-each>
                <xsl:for-each select="w:r/w:rPr/w:rStyle[substring(@w:val,1,10)='Bodymatter']">
                    <xsl:if test="d:PushPageStyle($myObj,@w:val)"/>
                </xsl:for-each>
                <xsl:for-each select="w:r/w:rPr/w:rStyle[substring(@w:val,1,10)='Rearmatter']">
                    <xsl:if test="d:PushPageStyle($myObj,@w:val)"/>
                </xsl:for-each>
                <xsl:sequence select="d:IncrementCheckingParagraph($myObj)"/> <!-- empty -->
            </xsl:if>
        </xsl:for-each>

        <xsl:if test="d:IsInvalidPageStylesSequence($myObj)='true'">
            <xsl:message terminate="yes">
                <xsl:value-of select="d:GetPageStylesErrors($myObj)"/>
            </xsl:message>
        </xsl:if>

    </xsl:template>

    <xsl:template name="SetCurrentMatterType">
        <xsl:param name="isRecursiveCall" />
        <xsl:param name="nodePosition" />
        <xsl:message terminate="no">progress:setcurrentmattertype</xsl:message>
        <xsl:if test="name()='w:p'">
            <xsl:choose>
                <xsl:when test="(count(w:pPr/w:pStyle[substring(@w:val,1,11)='Frontmatter'])=1 or count(w:r/w:rPr/w:rStyle[substring(@w:val,1,11)='Frontmatter'])=1)">
                    <xsl:if test="d:SetCurrentMatterType($myObj,'Frontmatter')"/>
                </xsl:when>
                <xsl:when test="(count(w:pPr/w:pStyle[substring(@w:val,1,10)='Bodymatter'])=1 or count(w:r/w:rPr/w:rStyle[substring(@w:val,1,10)='Bodymatter'])=1)">
                    <xsl:if test="d:SetCurrentMatterType($myObj,'Bodymatter')"/>
                </xsl:when>
                <xsl:when test="(count(w:pPr/w:pStyle[substring(@w:val,1,10)='Rearmatter'])=1) or (count(w:r/w:rPr/w:rStyle[substring(@w:val,1,10)='Rearmatter'])=1)">
                    <xsl:if test="d:SetCurrentMatterType($myObj,'Rearmatter')"/>
                </xsl:when>
            </xsl:choose>
        </xsl:if>
    </xsl:template>

</xsl:stylesheet>
