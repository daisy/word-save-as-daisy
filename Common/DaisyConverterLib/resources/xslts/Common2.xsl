<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:v="urn:schemas-microsoft-com:vml"
	xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:myObj="urn:Daisy" exclude-result-prefixes="w pic wp dcterms xsi cp dc a r v dcmitype myObj xsl">
	<!--<xsl:param name="sOperators"/>
	<xsl:param name="sMinuses"/>
	<xsl:param name="sNumbers"/>
	<xsl:param name="sZeros"/>-->
	<xsl:output method="xml" indent="no" />

	<!--Template for adding Levels-->
	<xsl:template name="AddLevel">
		<!--Parameter levelValue holds the value of the current level -->
		<xsl:param name="levelValue"/>
		<!--Parameter check holds the value that checks for different level values-->
		<xsl:param name="check"/>
		<xsl:param name="verhead"/>
		<xsl:param name="custom"/>
		<xsl:param name="mastersubhead"/>
		<xsl:param name="abValue"/>
		<xsl:param name="txt"/>
		<xsl:param name="lvlcharStyle"/>
		<xsl:param name="sOperators"/>
		<xsl:param name="sMinuses"/>
		<xsl:param name="sNumbers"/>
		<xsl:param name="sZeros"/>
		<xsl:message terminate="no">progress:addlevel</xsl:message>
		<!--Pushing level into the stack-->
		<xsl:variable name ="headingIncrementCounters" select="myObj:IncrementHeadingCounters($levelValue,substring-after($txt,'!'),$abValue)"/>
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:variable name ="copyCounter" select="myObj:CopyToBaseCounter(substring-after($txt,'!'))"/>
        <xsl:message terminate="no">progress:parahandler</xsl:message>
		
        <xsl:variable name="level" select="myObj:PushLevel($levelValue)"/>
        
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:choose>
			<!--Checking the level value-->
			<xsl:when test="$level &lt; 7">
				<!--Levels upto 6-->
				<!--Creating level element-->
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','level',$level,'&gt;')"/>
				<xsl:if test="(myObj:GetPageNum()=0) and (myObj:CheckTocOccur()=1) and ($custom='Automatic')">
					<xsl:if test="preceding-sibling::w:sdt[1]/w:sdtPr/w:docPartObj/w:docPartGallery/@w:val='Cover Pages'">
						<xsl:variable name="increment" select="myObj:IncrementPage()"/>
					</xsl:if>
					<xsl:if test="not(w:r[1]/w:br/@w:type='page') or not(w:r[1]/w:lastRenderedPageBreak)">
						<xsl:call-template name="DefaultPageNum"/>
					</xsl:if>
				</xsl:if>
				<!--Checking whether heading is present in the document-->
				<xsl:if test="$check!=0">
					<!--checking for custom style set for page numbers-->
					<xsl:if test="$custom='Automatic'">
						<xsl:choose>
							<!--Checking page breaks-->
							<xsl:when test="(w:r/w:br/@w:type='page') or (w:r/w:lastRenderedPageBreak)">
								<xsl:choose>
									<xsl:when test="((w:r/w:br/@w:type='page') and not((following-sibling::w:p[1]/w:pPr/w:sectPr) or (following-sibling::w:p[2]/w:r/w:lastRenderedPageBreak) or (following-sibling::w:p[1]/w:r/w:lastRenderedPageBreak) or (following-sibling::w:sdt[1]/w:sdtPr/w:docPartObj/w:docPartGallery/@w:val='Table of Contents')))">
										<xsl:variable name="increment" select="myObj:IncrementPage()"/>
										<xsl:call-template name="SectionBreak">
											<xsl:with-param name="count" select="'1'"/>
											<xsl:with-param name="node" select="'Para'"/>
										</xsl:call-template>
									</xsl:when>
									<xsl:when test="(w:r/w:lastRenderedPageBreak)">
										<xsl:variable name="increment" select="myObj:IncrementPage()"/>
										<xsl:call-template name="SectionBreak">
											<xsl:with-param name="count" select="'1'"/>
											<xsl:with-param name="node" select="'Para'"/>
										</xsl:call-template>
									</xsl:when>
								</xsl:choose>
							</xsl:when>
						</xsl:choose>
					</xsl:if>
					<!--Calling tmpHeading template for adding Levels-->

					<!-- DB :  Check if PageNumberDAISY style is applied to skip heading styles in output file when this style is applied.  -->
					<xsl:variable name="IsPageNumberDAISYApplied">
						<xsl:choose>
							<xsl:when test="w:pPr/w:rPr/w:rStyle/@w:val='PageNumberDAISY'">
								<xsl:value-of select="'true'"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="'false'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:variable>

					<!-- DB : write header in output when PageNumberDAISY is not applied  -->
					<xsl:if test="not($IsPageNumberDAISYApplied='true')">
						<xsl:call-template name="tmpHeading">
							<xsl:with-param name="level" select="$level"/>
						</xsl:call-template>
					</xsl:if>

					<!--Calling ParaHandler template for heading text-->
					<xsl:call-template name="TempLevelSpan">
						<xsl:with-param name ="verhead" select="$verhead"/>
						<xsl:with-param name ="custom" select="$custom"/>
						<xsl:with-param name="level" select="$level"/>
						<xsl:with-param name ="txt" select="myObj:TextHeading(substring-before($txt,'|'),substring-before(substring-after($txt,'|'),'!'),substring-after($txt,'!'),$level)"/>
						<xsl:with-param name ="mastersubhead" select="$mastersubhead"/>
						<xsl:with-param name="lvlcharStyle" select="$lvlcharStyle"/>
						<xsl:with-param name ="sOperators" select="$sOperators"/>
						<xsl:with-param name ="sMinuses" select="$sMinuses"/>
						<xsl:with-param name ="sNumbers" select="$sNumbers"/>
						<xsl:with-param name ="sZeros" select="$sZeros"/>
					</xsl:call-template>

					<!-- DB : write header in output when PageNumberDAISY is not applied  -->
					<xsl:if test="not($IsPageNumberDAISYApplied='true')">
						<!--Calling tmpAbbrAcrHeading template setting AbbrAcr flag and closing heading tag-->
						<xsl:call-template name="tmpAbbrAcrHeading">
							<xsl:with-param name="level" select="$level"/>
							<xsl:with-param name="levelValue" select="$levelValue"/>
						</xsl:call-template>
					</xsl:if>

				</xsl:if>
			</xsl:when>
			<!--Levels above 6-->
			<xsl:when test="$level &gt; 6">
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','level',6,'&gt;')"/>
				<!--Calling tmpHeading template for adding Levels-->
				<xsl:call-template name="tmpHeading">
					<xsl:with-param name="level" select="'6'"/>
				</xsl:call-template>
				<!--Calling ParaHandler template for heading text-->
				<xsl:call-template name="TempLevelSpan">
					<xsl:with-param name ="verhead" select="$verhead"/>
					<xsl:with-param name ="custom" select="$custom"/>
					<xsl:with-param name="level" select="$level"/>
					<xsl:with-param name ="txt" select="myObj:TextHeading(substring-before($txt,'|'),substring-before(substring-after($txt,'|'),'!'),substring-after($txt,'!'),$level)"/>
					<xsl:with-param name ="mastersubhead" select="$mastersubhead"/>
					<xsl:with-param name="lvlcharStyle" select="$lvlcharStyle"/>
					<xsl:with-param name ="sOperators" select="$sOperators"/>
					<xsl:with-param name ="sMinuses" select="$sMinuses"/>
					<xsl:with-param name ="sNumbers" select="$sNumbers"/>
					<xsl:with-param name ="sZeros" select="$sZeros"/>
				</xsl:call-template>


				<!--Calling tmpAbbrAcrHeading template setting AbbrAcr flag and closing heading tag-->
				<xsl:call-template name="tmpAbbrAcrHeading">
					<xsl:with-param name="level" select="$level"/>
					<xsl:with-param name="levelValue" select="$levelValue"/>
				</xsl:call-template>
			</xsl:when>
		</xsl:choose>
	</xsl:template>

	<xsl:template name ="TempLevelSpan">
		<xsl:param name="verhead"/>
		<xsl:param name ="custom"/>
		<xsl:param name="level"/>
		<xsl:param name ="txt"/>
		<xsl:param name ="mastersubhead"/>
		<xsl:param name="lvlcharStyle"/>
		<xsl:param name="sOperators"/>
		<xsl:param name="sMinuses"/>
		<xsl:param name="sNumbers"/>
		<xsl:param name="sZeros"/>
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:choose>
			<xsl:when test="$lvlcharStyle='True'">
				<xsl:choose>
					<xsl:when test="w:pPr/w:ind[@w:left] and w:pPr/w:ind[@w:right]">
						<xsl:variable name="val" select="w:pPr/w:ind/@w:left"/>
						<xsl:variable name="val_left" select="($val div 1440)"/>
						<xsl:variable name="valright" select="w:pPr/w:ind/@w:right"/>
						<xsl:variable name="val_right" select="($valright div 1440)"/>
						<span class="{concat('text-indent:', 'right=',$val_right,'in',';left=',$val_left,'in')}">
							<xsl:call-template name="ParaHandler">
								<xsl:with-param name="flag" select="'0'"/>
								<xsl:with-param name="VERSION" select="$verhead"/>
								<xsl:with-param name ="custom" select="$custom"/>
								<xsl:with-param name="level" select="$level"/>
								<xsl:with-param name ="txt" select="$txt"/>
								<xsl:with-param name="sOperators" select="$sOperators"/>
								<xsl:with-param name="sMinuses" select="$sMinuses"/>
								<xsl:with-param name="sNumbers" select="$sNumbers"/>
								<xsl:with-param name="sZeros" select="$sZeros"/>
								<xsl:with-param name ="mastersubpara" select="$mastersubhead"/>
								<xsl:with-param name="charparahandlerStyle" select="$lvlcharStyle"/>
							</xsl:call-template>
						</span>
					</xsl:when>
					<xsl:when test="w:pPr/w:ind[@w:left] and  w:pPr/w:jc">
						<xsl:variable name="val" select="w:pPr/w:ind/@w:left"/>
						<xsl:variable name="val_left" select="($val div 1440)"/>
						<xsl:variable name="val1" select="w:pPr/w:jc/@w:val"/>
						<span class="{concat('text-indent:',';left=',$val_left,'in',';text-align:',$val1)}">
							<xsl:call-template name="ParaHandler">
								<xsl:with-param name="flag" select="'0'"/>
								<xsl:with-param name="VERSION" select="$verhead"/>
								<xsl:with-param name ="custom" select="$custom"/>
								<xsl:with-param name="level" select="$level"/>
								<xsl:with-param name ="txt" select="$txt"/>
								<xsl:with-param name="sOperators" select="$sOperators"/>
								<xsl:with-param name="sMinuses" select="$sMinuses"/>
								<xsl:with-param name="sNumbers" select="$sNumbers"/>
								<xsl:with-param name="sZeros" select="$sZeros"/>
								<xsl:with-param name ="mastersubpara" select="$mastersubhead"/>
								<xsl:with-param name="charparahandlerStyle" select="$lvlcharStyle"/>
							</xsl:call-template>
						</span>
					</xsl:when>
					<xsl:when test="w:pPr/w:ind[@w:left]">
						<xsl:variable name="val" select="w:pPr/w:ind/@w:left"/>
						<xsl:variable name="val_left" select="($val div 1440)"/>
						<span class="{concat('text-indent:',$val_left,'in')}">
							<xsl:call-template name="ParaHandler">
								<xsl:with-param name="flag" select="'0'"/>
								<xsl:with-param name="VERSION" select="$verhead"/>
								<xsl:with-param name ="custom" select="$custom"/>
								<xsl:with-param name="level" select="$level"/>
								<xsl:with-param name ="txt" select="$txt"/>
								<xsl:with-param name="sOperators" select="$sOperators"/>
								<xsl:with-param name="sMinuses" select="$sMinuses"/>
								<xsl:with-param name="sNumbers" select="$sNumbers"/>
								<xsl:with-param name="sZeros" select="$sZeros"/>
								<xsl:with-param name ="mastersubpara" select="$mastersubhead"/>
								<xsl:with-param name="charparahandlerStyle" select="$lvlcharStyle"/>
							</xsl:call-template>
						</span>

					</xsl:when>
					<xsl:when test="w:pPr/w:ind[@w:right]">
						<xsl:variable name="val" select="w:pPr/w:ind/@w:right"/>
						<xsl:variable name="val_right" select="($val div 1440)"/>
						<span class="{concat('text-indent:',$val_right,'in')}">
							<xsl:call-template name="ParaHandler">
								<xsl:with-param name="flag" select="'0'"/>
								<xsl:with-param name="VERSION" select="$verhead"/>
								<xsl:with-param name ="custom" select="$custom"/>
								<xsl:with-param name="level" select="$level"/>
								<xsl:with-param name ="txt" select="$txt"/>
								<xsl:with-param name="sOperators" select="$sOperators"/>
								<xsl:with-param name="sMinuses" select="$sMinuses"/>
								<xsl:with-param name="sNumbers" select="$sNumbers"/>
								<xsl:with-param name="sZeros" select="$sZeros"/>
								<xsl:with-param name ="mastersubpara" select="$mastersubhead"/>
								<xsl:with-param name="charparahandlerStyle" select="$lvlcharStyle"/>
							</xsl:call-template>
						</span>

					</xsl:when>
					<xsl:when test="w:pPr/w:jc">
						<xsl:variable name="val" select="w:pPr/w:jc/@w:val"/>
						<span class="{concat('text-align:',$val)}">
							<xsl:call-template name="ParaHandler">
								<xsl:with-param name="flag" select="'0'"/>
								<xsl:with-param name="VERSION" select="$verhead"/>
								<xsl:with-param name ="custom" select="$custom"/>
								<xsl:with-param name="level" select="$level"/>
								<xsl:with-param name ="txt" select="$txt"/>
								<xsl:with-param name="sOperators" select="$sOperators"/>
								<xsl:with-param name="sMinuses" select="$sMinuses"/>
								<xsl:with-param name="sNumbers" select="$sNumbers"/>
								<xsl:with-param name="sZeros" select="$sZeros"/>
								<xsl:with-param name ="mastersubpara" select="$mastersubhead"/>
								<xsl:with-param name="charparahandlerStyle" select="$lvlcharStyle"/>
							</xsl:call-template>
						</span>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="ParaHandler">
							<xsl:with-param name="flag" select="'0'"/>
							<xsl:with-param name="VERSION" select="$verhead"/>
							<xsl:with-param name ="custom" select="$custom"/>
							<xsl:with-param name="level" select="$level"/>
							<xsl:with-param name ="txt" select="$txt"/>
							<xsl:with-param name="sOperators" select="$sOperators"/>
							<xsl:with-param name="sMinuses" select="$sMinuses"/>
							<xsl:with-param name="sNumbers" select="$sNumbers"/>
							<xsl:with-param name="sZeros" select="$sZeros"/>
							<xsl:with-param name ="mastersubpara" select="$mastersubhead"/>
							<xsl:with-param name="charparahandlerStyle" select="$lvlcharStyle"/>
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:call-template name="ParaHandler">
					<xsl:with-param name="flag" select="'0'"/>
					<xsl:with-param name="VERSION" select="$verhead"/>
					<xsl:with-param name ="custom" select="$custom"/>
					<xsl:with-param name="level" select="$level"/>
					<xsl:with-param name ="txt" select="$txt"/>
					<xsl:with-param name="sOperators" select="$sOperators"/>
					<xsl:with-param name="sMinuses" select="$sMinuses"/>
					<xsl:with-param name="sNumbers" select="$sNumbers"/>
					<xsl:with-param name="sZeros" select="$sZeros"/>
					<xsl:with-param name ="mastersubpara" select="$mastersubhead"/>
					<xsl:with-param name="charparahandlerStyle" select="$lvlcharStyle"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>

	</xsl:template>

	<!--Template to Close Levels-->
	<xsl:template name="CloseLevel">
		<xsl:param name="CurrentLevel"/>
		<xsl:message terminate="no">progress:closelevel</xsl:message>
		<!--<xsl:value-of select="$CurrentLevel"/>-->
		<!--Peeking the top value of the stack-->
		<xsl:variable name="PeekLevel" select="myObj:PeekLevel()"/>
		<!--<xsl:value-of select="$PeekLevel"/>-->
		<!--If top level is less than or equal to current level then PoP the Stack and close that level-->
		<xsl:choose>
			<xsl:when test ="$CurrentLevel &gt; 6 and $PeekLevel = 6 ">
				<xsl:variable name="PopLevel" select="myObj:PoPLevel()"/>
				<!--Close that level-->
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/level',$PopLevel,'&gt;')"/>
				<!--Loop through until we have closed all the Lower levels-->
			</xsl:when>
			<xsl:otherwise>
				<xsl:if test="$CurrentLevel &lt;=$PeekLevel and $PeekLevel !=0">
					<xsl:variable name="PopLevel" select="myObj:PoPLevel()"/>
					<!--Close that level-->
          <p/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/level',$PopLevel,'&gt;')"/>
					<!--Loop through until we have closed all the Lower levels-->
					<xsl:call-template name="CloseLevel">
						<xsl:with-param name="CurrentLevel" select="$CurrentLevel"/>
					</xsl:call-template>
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

  <!--Insert page break in list-->
  <xsl:template name="PageInList">
    <xsl:param name ="custom"/>
    
    <xsl:for-each select ="./node()">
      <xsl:if test="name()='w:r'">
        <xsl:if test="((w:lastRenderedPageBreak) or (w:br/@w:type='page'))
                      and ($custom='Automatic') ">
          <xsl:choose>
            <xsl:when test="not(w:t) and (w:lastRenderedPageBreak) and (w:br/@w:type='page')">
              <xsl:if test="not(../following-sibling::w:sdt[1]/w:sdtPr/w:docPartObj/w:docPartGallery/@w:val='Table of Contents')">
                <xsl:if test="not(../preceding-sibling::node()[1]/w:pPr/w:sectPr)">
                  <xsl:variable name="increment" select="myObj:IncrementPage()"/>
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
                </xsl:if>
              </xsl:if>
            </xsl:when>
            <!--Checking for page breaks and populating page numbers.-->
            <xsl:when test="( (w:br/@w:type='page')
                            and not((../following-sibling::w:p[1]/w:pPr/w:sectPr)
                                or (../following-sibling::w:p[2]/w:r/w:lastRenderedPageBreak)
                                or (../following-sibling::w:p[1]/w:r/w:lastRenderedPageBreak)
                                or (../following-sibling::w:sdt[1]/w:sdtPr/w:docPartObj/w:docPartGallery/@w:val='Table of Contents')) )">
              <!--Incrementing page numbers-->
              <xsl:variable name="increment" select="myObj:IncrementPage()"/>
              <!--calling template to initialize page number information-->
              <xsl:call-template name="SectionBreak">
                <xsl:with-param name="count" select="'1'"/>
                <xsl:with-param name="node" select="'body'"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="(w:lastRenderedPageBreak)
                            and not(../w:pPr/w:sectPr
                                or ../following-sibling::w:sdt[1]/w:sdtPr/w:docPartObj/w:docPartGallery/@w:val='Table of Contents')">
              <xsl:variable name="increment" select="myObj:IncrementPage()"/>
              <!--calling template to initialize page number information-->
              <xsl:call-template name="SectionBreak">
                <xsl:with-param name="count" select="'1'"/>
                <xsl:with-param name="node" select="'body'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:if>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

	<!--Template to Add List-->
	<xsl:template name="addlist">
		<xsl:param name="openId"/>
		<xsl:param name="openlvl"/>
		<xsl:param name ="verlist"/>
		<xsl:param name ="custom"/>
		<xsl:param name="numFmt"/>
		<xsl:param name="lText"/>
		<xsl:param name="lstcharStyle"/>
		<xsl:message terminate="no">progress:addlist</xsl:message>
		<!--Pushes the current level into the stack-->
		<xsl:variable name="PeekLevel" select="myObj:ListPeekLevel()"/>

		<!--Checking the current level with the PeekLevel in the stack-->
		<xsl:if test="$PeekLevel=$openlvl">
			<xsl:variable name="PeekLevel1" select="myObj:ListPush($openlvl)"/>
			<xsl:variable name ="IncListCounter" select="myObj:IncrementListCounters($openlvl,$openId)"/>
			<xsl:variable name ="txt" select="myObj:TextList($numFmt,$lText,$openId,$openlvl)"/>
			<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','li','&gt;')"/>

			<xsl:choose>
				<xsl:when  test="($numFmt='lowerLetter')
                      or ($numFmt='lowerRoman')
                      or ($numFmt='upperRoman')
                      or ($numFmt='upperLetter')
                      or ($numFmt='decimalZero')">
					<!--<xsl:value-of select="$txt"/>-->
					<xsl:call-template name="ParagraphStyle">
						<xsl:with-param name ="custom" select="$custom"/>
						<xsl:with-param name="txt" select="$txt"/>
						<xsl:with-param name="characterparaStyle" select="$lstcharStyle"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<xsl:call-template name="ParagraphStyle">
						<xsl:with-param name ="custom" select="$custom"/>
						<!--<xsl:with-param name="txt" select="$txt"/>-->
						<xsl:with-param name="characterparaStyle" select="$lstcharStyle"/>
					</xsl:call-template>
				</xsl:otherwise>
			</xsl:choose>

		</xsl:if>
    
		<!--Checking the current level with the PeekLevel in the stack-->
		<xsl:if test="$openlvl &gt; $PeekLevel">
			<xsl:variable name="diffLevel" select="myObj:DiffLevel($openlvl,$PeekLevel)"/>
			<xsl:choose>
				<xsl:when test="$diffLevel = 1">
					<xsl:variable name="PeekLevel1" select="myObj:ListPush($openlvl)"/>
					<xsl:variable name ="IncListCounter" select="myObj:IncrementListCounters($openlvl,$openId)"/>
					<xsl:variable name ="txt" select="myObj:TextList($numFmt,$lText,$openId,$openlvl)"/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','li','&gt;')"/>

					<xsl:choose>
						<xsl:when  test="($numFmt='lowerLetter')
                          or ($numFmt='lowerRoman')
                          or ($numFmt='upperRoman')
                          or ($numFmt='upperLetter')
                          or ($numFmt='decimalZero')">
							<xsl:call-template name="ParagraphStyle">
								<xsl:with-param name ="custom" select="$custom"/>
								<xsl:with-param name="txt" select="$txt"/>
								<xsl:with-param name="characterparaStyle" select="$lstcharStyle"/>
							</xsl:call-template>
						</xsl:when>
						<xsl:otherwise>
							<xsl:call-template name="ParagraphStyle">
								<xsl:with-param name ="custom" select="$custom"/>
								<!--<xsl:with-param name="txt" select="$txt"/>-->
								<xsl:with-param name="characterparaStyle" select="$lstcharStyle"/>
							</xsl:call-template>
						</xsl:otherwise>
					</xsl:choose>

				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','li','&gt;')"/>
					<xsl:variable name="reduceOne" select="myObj:ReduceOne($diffLevel)"/>
					<xsl:call-template name="recursive">
						<xsl:with-param name="rec" select="$reduceOne"/>
					</xsl:call-template>
					<xsl:value-of select="myObj:Increment($openlvl,$PeekLevel,$openId)"/>
					<xsl:variable name ="txt" select="myObj:TextList($numFmt,$lText,$openId,$openlvl)"/>

					<xsl:choose>
						<xsl:when  test="($numFmt='lowerLetter')
                          or ($numFmt='lowerRoman')
                          or ($numFmt='upperRoman')
                          or ($numFmt='upperLetter')
                          or ($numFmt='decimalZero')">
							<xsl:call-template name="ParagraphStyle">
								<xsl:with-param name ="custom" select="$custom"/>
								<xsl:with-param name="txt" select="$txt"/>
								<xsl:with-param name="characterparaStyle" select="$lstcharStyle"/>
							</xsl:call-template>
						</xsl:when>
						<xsl:otherwise>
							<xsl:call-template name="ParagraphStyle">
								<xsl:with-param name ="custom" select="$custom"/>
								<!--<xsl:with-param name="txt" select="$txt"/>-->
								<xsl:with-param name="characterparaStyle" select="$lstcharStyle"/>
							</xsl:call-template>
						</xsl:otherwise>
					</xsl:choose>

				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>

	</xsl:template>

	<!--Template to close the List-->
	<xsl:template name="closelist">
		<xsl:param name="close"/>
		<xsl:message terminate="no">progress:closelist</xsl:message>
		<!--Gets the current level of the stack  -->
		<xsl:variable name="PeekLevel" select="myObj:ListPeekLevel()"/>
		<!--Checking the current level with the PeekLevel in the stack-->
		<xsl:if test="$close &lt; $PeekLevel">
			<!--PoPs one level from the stack-->
			<xsl:variable name="PopLevel" select="myObj:ListPoPLevel()"/>
      <xsl:if test="not(preceding-sibling::node()[1][w:r/w:rPr/w:rStyle[substring(@w:val,1,15)='PageNumberDAISY']])">
			  <xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/li','&gt;')"/>
      </xsl:if>
			<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/list','&gt;')"/>
			<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/li','&gt;')"/>
			<!--Loop through until we have closed all the Lower levels-->
			<xsl:call-template name="closelist">
				<xsl:with-param name="close" select="$close"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<xsl:template name="recursive">
		<xsl:param name="rec"/>
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:variable name="aquote">"</xsl:variable>
		<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','list ','type=',$aquote,'ol',$aquote,'&gt;')"/>
		<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','li','&gt;')"/>
		<xsl:variable name="dec" select="myObj:Decrement($rec)"/>
		<xsl:if test="$dec!=0">
			<xsl:call-template name="recursive">
				<xsl:with-param name="rec" select="$dec"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<xsl:template name="recStart">
		<xsl:param name="abstLevel"/>
		<xsl:param name="level"/>
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:choose>
			<xsl:when test="$level='0'">
				<xsl:variable name="strStart">
					<xsl:value-of select="document('word/numbering.xml')//w:numbering/w:abstractNum[@w:abstractNumId=$abstLevel]/w:lvl[@w:ilvl=$level]/w:start/@w:val"/>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test="$strStart=''">
						<xsl:variable name="appendString" select="myObj:StartString($level,'0')"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:variable name="appendString" select="myObj:StartString($level,$strStart)"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="strStart">
					<xsl:value-of select="document('word/numbering.xml')//w:numbering/w:abstractNum[@w:abstractNumId=$abstLevel]/w:lvl[@w:ilvl=$level]/w:start/@w:val"/>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test="$strStart=''">
						<xsl:variable name="appendString" select="myObj:StartString($level,'0')"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:variable name="appendString" select="myObj:StartString($level,$strStart)"/>
					</xsl:otherwise>
				</xsl:choose>
				<xsl:variable name="dec" select="myObj:DecrementStart($level)"/>
				<xsl:call-template name="recStart">
					<xsl:with-param name="abstLevel" select="$abstLevel"/>
					<xsl:with-param name="level" select="$dec"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!--Template to Close Complex List-->
	<xsl:template name="ComplexListClose">
		<xsl:param name="close"/>
		<xsl:message terminate="no">progress:closelist</xsl:message>
		<!--Gets the current level of the stack  -->
		<xsl:variable name="PeekLevel" select="myObj:ListPeekLevel()"/>
		<!--Checking the current level with the PeekLevel in the stack-->
		<xsl:if test="$close &lt; $PeekLevel">
			<!--PoPs one level from the stack-->
			<xsl:variable name="PoPLevel" select="myObj:ListPoPLevel()"/>
			<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/list','&gt;')"/>
			<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/li','&gt;')"/>
			<!--Loop through until we have closed all the Lower levels-->
			<xsl:call-template name="ComplexListClose">
				<xsl:with-param name="close" select="$close"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<!--Template to Close All nested List-->
	<xsl:template name="CloseLastlist">
		<xsl:param name="close"/>
    <xsl:param name="custom"/>
		<xsl:message terminate="no">progress:closelist</xsl:message>
		<!--Gets the current level of the stack  -->
		<xsl:variable name="PeekLevel" select="myObj:ListPeekLevel()"/>
		<!--Checking the current level with the PeekLevel in the stack-->
		<xsl:if test="$close &lt;= $PeekLevel">
			<!--PoPs one level from the stack-->
			<xsl:variable name="PoPLevel" select="myObj:ListPoPLevel()"/>
			<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/li','&gt;')"/>
			<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/list','&gt;')"/>
			<xsl:if test="$PeekLevel!=0">
				<!--Loop through until we have closed all the Lower levels-->
				<xsl:call-template name="CloseLastlist">
					<xsl:with-param name="close" select="$close"/>
          <xsl:with-param name="custom" select="$custom"/>
				</xsl:call-template>
			</xsl:if>
		</xsl:if>
	</xsl:template>

	<!--Template for default Page number-->
	<xsl:template name="DefaultPageNum">
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:if test="myObj:ReturnPageNum()&lt;=1">
			<xsl:variable name="increment" select="myObj:IncrementPage()"/>
		</xsl:if>
		<xsl:if test="preceding-sibling::node()[1]/w:pPr/w:sectPr">
			<xsl:variable name="setConPageBreak" select="myObj:SetConPageBreak()"/>
		</xsl:if>
		<!--Traversing through each node-->
		<xsl:for-each select="following-sibling::node()">
			<xsl:choose>
				<!--Checking for paragraph section break-->
				<xsl:when test="w:pPr/w:sectPr">
					<xsl:if test="myObj:CheckSection()=1">
						<xsl:variable name="sectionInfo" select="myObj:SectionCounter(w:pPr/w:sectPr/w:pgNumType/@w:fmt,w:pPr/w:sectPr/w:pgNumType/@w:start)"/>
						<xsl:choose>
							<!--Checking if page start and page format is present-->
							<xsl:when test="(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
								<!--Calling tempale for page number text-->
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
									<xsl:with-param name="matter" select="'body'"/>
									<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
								</xsl:call-template>
							</xsl:when>
							<!--Checking if page format is present and not page start-->
							<xsl:when test="(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and not(w:pPr/w:sectPr/w:pgNumType/@w:start)">
								<!--Calling tempale for page number text-->
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
									<xsl:with-param name="matter" select="'body'"/>
									<xsl:with-param name="counter" select="'0'"/>
								</xsl:call-template>
							</xsl:when>
							<!--Checking if page start is present and not page format-->
							<xsl:when test="not(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
								<!--Calling tempale for page number text-->
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
									<xsl:with-param name="matter" select="'body'"/>
									<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
								</xsl:call-template>
							</xsl:when>
							<xsl:otherwise>
								<!--If both are not present-->
								<!--Calling tempale for page number text-->
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
									<xsl:with-param name="matter" select="'body'"/>
									<xsl:with-param name="counter" select="'0'"/>
								</xsl:call-template>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:if>
				</xsl:when>
				<!--Checking for Section in a document-->
				<xsl:when test="name()='w:sectPr'">
					<xsl:if test="myObj:CheckSection()=1">
						<xsl:variable name="sectionInfo" select="myObj:SectionCounter(w:pgNumType/@w:fmt,w:pgNumType/@w:start)"/>
						<xsl:choose>
							<!--Checking if page start and page format is present-->
							<xsl:when test="(w:pgNumType/@w:fmt) and (w:pgNumType/@w:start)">
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="w:pgNumType/@w:fmt"/>
									<xsl:with-param name="matter" select="'body'"/>
									<xsl:with-param name="counter" select="w:pgNumType/@w:start"/>
								</xsl:call-template>
							</xsl:when>
							<!--Checking if page format is present and not page start-->
							<xsl:when test="(w:pgNumType/@w:fmt) and not(w:pgNumType/@w:start)">
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="w:pgNumType/@w:fmt"/>
									<xsl:with-param name="matter" select="'body'"/>
									<xsl:with-param name="counter" select="'0'"/>
								</xsl:call-template>
							</xsl:when>
							<!--Checking if page start is present and not page format-->
							<xsl:when test="not(w:pgNumType/@w:fmt) and (w:pgNumType/@w:start)">
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
									<xsl:with-param name="matter" select="'body'"/>
									<xsl:with-param name="counter" select="w:pgNumType/@w:start"/>
								</xsl:call-template>
							</xsl:when>
							<xsl:otherwise>
								<!--If both are not present-->
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
									<xsl:with-param name="matter" select="'body'"/>
									<xsl:with-param name="counter" select="'0'"/>
								</xsl:call-template>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:if>
				</xsl:when>
			</xsl:choose>
		</xsl:for-each>
		<xsl:variable name="reSetConPageBreak" select="myObj:ResetSetConPageBreak()"/>
		<xsl:variable name="initialize" select="myObj:InitalizeCheckSection()"/>
	</xsl:template>

	<xsl:template name="tmpHeading">
		<xsl:param name="level"/>
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:choose>
			<xsl:when test="(w:r/w:rPr/w:lang) or (w:r/w:rPr/w:rFonts/@w:hint)">
				<xsl:call-template name="LanguagesPara">
					<xsl:with-param name="Attribute" select="'1'"/>
					<xsl:with-param name ="level" select="concat('h',$level)"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',concat('h',$level),'&gt;')"/>
			</xsl:otherwise>
		</xsl:choose>

	</xsl:template>

	<xsl:template name="tmpAbbrAcrHeading">
		<xsl:param name="level"/>
		<xsl:param name="levelValue"/>
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:choose>
			<xsl:when test="(myObj:AbbrAcrFlag()='1') and not(w:r/w:pict/v:shape/v:textbox)">
				<xsl:variable name="temp">
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',concat('/h',$level),'&gt;')"/>
				</xsl:variable>
				<xsl:variable name="abbrpara" select="myObj:PushAbrAcrhead($temp)"/>
				<xsl:if test="myObj:ListMasterSubFlag()='1'">
					<xsl:variable name ="curLevel" select="myObj:PeekLevel()"/>
					<xsl:value-of disable-output-escaping="yes" select="myObj:ClosingMasterSub($curLevel)"/>
					<xsl:value-of disable-output-escaping="yes" select="myObj:PeekMasterSubdoc()"/>
					<xsl:variable name="masterSubReSet" select="myObj:MasterSubResetFlag()"/>
					<xsl:value-of disable-output-escaping="yes" select="myObj:OpenMasterSub($curLevel)"/>
				</xsl:if>
			</xsl:when>
			<xsl:otherwise>
				<xsl:if test="not(w:r/w:pict/v:shape/v:textbox)">
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',concat('/h',$level),'&gt;')"/>
					<xsl:if test="myObj:ListMasterSubFlag()='1'">
						<xsl:variable name ="curLevel" select="myObj:PeekLevel()"/>
						<xsl:value-of disable-output-escaping="yes" select="myObj:ClosingMasterSub($curLevel)"/>
						<xsl:value-of disable-output-escaping="yes" select="myObj:PeekMasterSubdoc()"/>
						<xsl:variable name="masterSubReSet" select="myObj:MasterSubResetFlag()"/>
						<xsl:value-of disable-output-escaping="yes" select="myObj:OpenMasterSub($curLevel)"/>
					</xsl:if>
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
		<xsl:if test="(following-sibling::node()[1][name()='w:sectPr']) or (following-sibling::node()[1]/w:pPr/w:pStyle[substring(@w:val,8,1)=$levelValue])">
			<p></p>
		</xsl:if>
	</xsl:template>

	<!--Template for Implementing Table-->
	<xsl:template name="TableHandler">
		<xsl:param name="parmVerTable"/>
		<xsl:param name="custom"/>
		<xsl:param name="mastersubtbl"/>
		<xsl:param name="characterStyle"/>
		<xsl:variable name="quote">"</xsl:variable>
		<xsl:message terminate="no">progress:tablehandler</xsl:message>
		<xsl:if test="$custom='Automatic'">
			<xsl:for-each select="w:tr/w:tc">
				<xsl:message terminate="no">progress:parahandler</xsl:message>
				<xsl:if test="((w:p/w:r/w:lastRenderedPageBreak) or (w:p/w:r/w:br/@w:type='page'))">
					<xsl:if test="not((../preceding-sibling::w:tr[1]/w:tc/w:p/w:r/w:lastRenderedPageBreak) or (../preceding-sibling::w:tr[1]/w:tc/w:p/w:r/w:br/@w:type='page') or (preceding-sibling::w:tc[1]/w:p/w:r/w:lastRenderedPageBreak) or (preceding-sibling::w:tc[1]/w:p/w:r/w:br/@w:type='page'))">
						<xsl:variable name="increment" select="myObj:IncrementPage()"/>
						<!--Calling SectionBreak template for getting the section page type information-->
						<xsl:call-template name="SectionBreak">
							<xsl:with-param name="count" select="'1'"/>
							<xsl:with-param name="node" select="'Table'"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:if>
			</xsl:for-each>
		</xsl:if>
		<table border="1">
			<!-- check previous sibling is a /w:pPr/w:pStyle[@w:val='Caption'], if so print w:t-->
			<xsl:if test="(preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Caption') or (preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Table-CaptionDAISY') or (following-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Table-CaptionDAISY')">
				<caption>
					<xsl:if test="(preceding-sibling::node()[1]/w:r/w:rPr/w:lang) or (preceding-sibling::node()[1]/w:r/w:rPr/w:rFonts/@w:hint)">
						<xsl:attribute name="xml:lang">
							<!--Calling PictureLanguage template for implementing language for the text in the Table-->
							<xsl:call-template name="PictureLanguage">
								<xsl:with-param name="CheckLang" select="'Table'"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="(preceding-sibling::node()[1]/w:r/w:rPr/w:rtl) or (preceding-sibling::node()[1]/w:pPr/w:bidi)">
						<xsl:variable name="varBdo">
							<xsl:call-template name="PictureLanguage">
								<xsl:with-param name="CheckLang" select="'Table'"/>
							</xsl:call-template>
						</xsl:variable>
						<xsl:variable name="bdoflag" select="myObj:SetcaptionFlag()"/>
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p  ','xml:lang=',$quote,$varBdo,$quote,'&gt;')"/>
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$varBdo,$quote,'&gt;')"/>
					</xsl:if>

					<xsl:if test="(preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Caption')">
						<xsl:for-each select="preceding-sibling::node()[1]/node()">
							<xsl:message terminate="no">progress:parahandler</xsl:message>
							<xsl:if test="name()='w:r'">
								<xsl:for-each select=".">
									<xsl:message terminate="no">progress:parahandler</xsl:message>
									<xsl:choose>
										<xsl:when test="w:noBreakHyphen">
											<xsl:text>-</xsl:text>
										</xsl:when>
										<xsl:otherwise>
											<xsl:call-template name ="TempCharacterStyle">
												<xsl:with-param name ="characterStyle" select="$characterStyle"/>
											</xsl:call-template>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:for-each>
							</xsl:if>
							<xsl:if test="name()='w:fldSimple'">
								<xsl:for-each select=".">
									<xsl:message terminate="no">progress:parahandler</xsl:message>
									<xsl:value-of select="w:r/w:t"/>
								</xsl:for-each>
							</xsl:if>
						</xsl:for-each>
					</xsl:if>
					<xsl:if test="(preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Table-CaptionDAISY')">
						<xsl:for-each select="preceding-sibling::node()[1]/node()">
							<xsl:message terminate="no">progress:parahandler</xsl:message>
							<xsl:if test="name()='w:r'">
								<xsl:for-each select=".">
									<xsl:message terminate="no">progress:parahandler</xsl:message>
									<xsl:choose>
										<xsl:when test="w:noBreakHyphen">
											<xsl:text>-</xsl:text>
										</xsl:when>
										<xsl:otherwise>
											<xsl:call-template name ="TempCharacterStyle">
												<xsl:with-param name ="characterStyle" select="$characterStyle"/>
											</xsl:call-template>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:for-each>
							</xsl:if>
							<xsl:if test="name()='w:fldSimple'">
								<xsl:for-each select=".">
									<xsl:message terminate="no">progress:parahandler</xsl:message>
									<xsl:value-of select="w:r/w:t"/>
								</xsl:for-each>
							</xsl:if>
						</xsl:for-each>
					</xsl:if>
					<xsl:if test="(following-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Table-CaptionDAISY')">
						<xsl:for-each select="following-sibling::node()[1]/node()">
							<xsl:message terminate="no">progress:parahandler</xsl:message>
							<xsl:if test="name()='w:r'">
								<xsl:for-each select=".">
									<xsl:message terminate="no">progress:parahandler</xsl:message>
									<xsl:choose>
										<xsl:when test="w:noBreakHyphen">
											<xsl:text>-</xsl:text>
										</xsl:when>
										<xsl:otherwise>
											<xsl:call-template name ="TempCharacterStyle">
												<xsl:with-param name ="characterStyle" select="$characterStyle"/>
											</xsl:call-template>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:for-each>
							</xsl:if>
							<xsl:if test="name()='w:fldSimple'">
								<xsl:for-each select=".">
									<xsl:message terminate="no">progress:parahandler</xsl:message>
									<xsl:value-of select="w:r/w:t"/>
								</xsl:for-each>
							</xsl:if>
						</xsl:for-each>
					</xsl:if>
					<xsl:if test="myObj:SetcaptionFlag()=2">
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
					</xsl:if>
					<xsl:variable name="captionflag" select="myObj:reSetcaptionFlag()"/>
					<xsl:if test="(preceding-sibling::w:p[1]/w:r/w:rPr/w:rtl) or (preceding-sibling::w:p[1]/w:pPr/w:bidi)">
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
					</xsl:if>
				</caption>
			</xsl:if>
			<xsl:if test="(following-sibling::w:p[1]/w:pPr/w:pStyle/@w:val='Caption')">
				<xsl:message terminate="no">translation.oox2Daisy.TableCaption</xsl:message>
			</xsl:if>
			<!--Checking whether alingment of the cell is set-->
			<xsl:if test="w:tr/w:tc/w:p/w:pPr/w:jc">
				<!--Creating col element of the table for specifying alignment of the cell content-->
				<xsl:variable name="colvalue">
					<xsl:value-of select="w:tr/w:tc/w:p/w:pPr/w:jc/@w:val"/>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test="not($colvalue='left') and not($colvalue='center') and not($colvalue='right') and not($colvalue='justify')and not($colvalue='char')">
						<col align="left"/>
					</xsl:when>
					<xsl:otherwise>
						<col align="{$colvalue}"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:if>
			<!--Checking for table header-->
			<xsl:if test="w:tr/w:trPr/w:tblHeader">
				<thead>
					<!--Looping through each row of the table-->
					<xsl:for-each select="w:tr">
						<xsl:message terminate="no">progress:parahandler</xsl:message>
						<!--Checking for table header-->
						<xsl:if test="w:trPr/w:tblHeader">
							<!--Looping through each row of the table-->
							<tr>
								<!--Looping through each cell of the table-->
								<xsl:for-each select="w:tc/w:p">
									<xsl:message terminate="no">progress:parahandler</xsl:message>
									<th>
										<!--Assinging value as Table head-->
										<xsl:call-template name="ParagraphStyle">
											<xsl:with-param name ="custom" select="$custom"/>
											<xsl:with-param name ="VERSION" select="$parmVerTable"/>
											<xsl:with-param name ="characterparaStyle" select="$characterStyle"/>
										</xsl:call-template>
									</th>
								</xsl:for-each>
							</tr>
						</xsl:if>
					</xsl:for-each>
				</thead>
			</xsl:if>
			<!--Checking for Table-FooterDAISY custom table style-->
			<xsl:if test="(w:tblPr/w:tblStyle/@w:val='Table-FooterDAISY') or (w:tblPr/w:tblStyle/@w:val='Table-footerDAISY') ">
				<tfoot>
					<!--Looping through each row of the table-->
					<xsl:for-each select="w:tr">
						<xsl:message terminate="no">progress:parahandler</xsl:message>
						<xsl:if test="position()=last()">
							<tr>
								<!--Looping through each cell of the table-->
								<xsl:for-each select="w:tc">
									<xsl:message terminate="no">progress:parahandler</xsl:message>
									<xsl:if test="(w:p) and (not((following-sibling::w:p[1]/w:pPr/w:pStyle[@w:val='DefinitionDataDAISY']) or (following-sibling::w:p[1]/w:r/w:rPr/w:rStyle[@w:val='DefinitionTermDAISY'])))">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','th','&gt;')"/>
									</xsl:if>
									<xsl:for-each select="w:p">
										<xsl:message terminate="no">progress:parahandler</xsl:message>
										<xsl:call-template name="ParagraphStyle">
											<xsl:with-param name ="custom" select="$custom"/>
											<xsl:with-param name ="VERSION" select="$parmVerTable"/>
											<xsl:with-param name ="characterparaStyle" select="$characterStyle"/>
										</xsl:call-template>
									</xsl:for-each>
									<xsl:if test="(w:p) and (not((following-sibling::w:p[1]/w:pPr/w:pStyle[@w:val='DefinitionDataDAISY']) or (following-sibling::w:p[1]/w:r/w:rPr/w:rStyle[@w:val='DefinitionTermDAISY'])))">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/th','&gt;')"/>
									</xsl:if>
								</xsl:for-each>
							</tr>
						</xsl:if>
					</xsl:for-each>
				</tfoot>
			</xsl:if>
			<tbody>
				<!--Counting each paragraph element for counting the number of rows spaned-->
				<!--Looping through each row of the table-->
				<xsl:for-each select="w:tr">
					<xsl:message terminate="no">progress:parahandler</xsl:message>
					<!--Checking if the row is not header row-->
					<xsl:if test="not(w:trPr/w:tblHeader) and not(w:trPr/w:cnfStyle)">
						<tr>
							<!--Looping through each cell of the table-->
							<xsl:for-each select="w:tc">
								<xsl:message terminate="no">progress:parahandler</xsl:message>
								<xsl:if test="w:tcPr/w:vMerge[@w:val='restart']">
									<xsl:variable name="columnPosition" select="position()"/>
									<xsl:for-each select="../following-sibling::w:tr/w:tc[$columnPosition]/w:tcPr">
										<xsl:message terminate="no">progress:parahandler</xsl:message>
										<xsl:if test="myObj:ReturnFlagRowspan()=0">
											<xsl:if test="(w:vMerge) and not(w:vMerge/@w:val='restart')">
												<xsl:variable name="rowspan" select="myObj:Rowspan()"/>
											</xsl:if>
										</xsl:if>
										<xsl:if test="not(w:vMerge)">
											<xsl:variable name="getflagRowspan" select="myObj:GetFlagRowspan()"/>
										</xsl:if>
									</xsl:for-each>
									<xsl:variable name="setflagRowspan" select="myObj:SetFlagRowspan()"/>
								</xsl:if>
								<!--If paragraph element exists-->
								<xsl:if test="w:p">
									<xsl:choose>
										<!--Checking for both colspan and rowspan-->
										<xsl:when test="(w:tcPr/w:gridSpan) and (w:tcPr/w:vMerge[@w:val='restart'])">
											<!--If Column span property is set the assinging the value-->
											<xsl:variable name="colspan" select="w:tcPr/w:gridSpan/@w:val"/>
											<!--variale holds the value of number of Rows span-->
											<!--Creating td tag with rowspan and colspan attribute-->
											<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','td ','colspan=',$quote,$colspan,$quote,' ',' rowspan=',$quote,myObj:GetRowspan()+1,$quote,'&gt;')"/>
										</xsl:when>
										<!--Checking for colspan and not rowspan-->
										<xsl:when test="(w:tcPr/w:gridSpan) and not(w:tcPr/w:vMerge[@w:val='restart'])">
											<!--colspan variable holds colspan value-->
											<xsl:variable name="colspan" select="w:tcPr/w:gridSpan/@w:val"/>
											<!--Creating td tag with colspan attribute-->
											<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','td ','colspan=',$quote,$colspan,$quote,'&gt;')"/>
										</xsl:when>
										<!--Checking for rowspan and not colspan-->
										<xsl:when test="(w:tcPr/w:vMerge[@w:val='restart']) and not(w:tcPr/w:gridSpan)">
											<!--rowspan variable holds rowspan value-->
											<!--Creating td tag with rowspan attribute-->
											<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','td ','rowspan=',$quote,myObj:GetRowspan()+1,$quote,'&gt;')"/>
										</xsl:when>
										<xsl:otherwise >
											<!--Opening the td tag-->
											<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','td ','&gt;')"/>
										</xsl:otherwise>
									</xsl:choose>
									<xsl:variable name ="var_heading">
										<xsl:for-each select="document('word/styles.xml')//w:styles/w:style/w:name[@w:val='heading 1']">
											<xsl:message terminate="no">progress:parahandler</xsl:message>
											<xsl:value-of select="../@w:styleId"/>
										</xsl:for-each>
									</xsl:variable>
									<xsl:variable name="setRowspan" select="myObj:SetRowspan()"/>
									<xsl:for-each select="w:p">
										<xsl:message terminate="no">progress:parahandler</xsl:message>
										<!--Calling paragraph template whenever w:p element is encountered.-->
										<xsl:call-template name="StyleContainer">
											<xsl:with-param name="VERSION" select="$parmVerTable"/>
											<xsl:with-param name ="custom" select="$custom"/>
											<xsl:with-param name ="styleHeading" select="$var_heading"/>
											<xsl:with-param name ="mastersubstyle" select="$mastersubtbl"/>
											<xsl:with-param name ="characterStyle" select="$characterStyle"/>
										</xsl:call-template>
									</xsl:for-each>
									<!--Checking for nested table-->
									<xsl:for-each select="child::w:tbl">
										<xsl:message terminate="no">progress:parahandler</xsl:message>
										<!--Calling template Tablehandler for nested tables-->
										<xsl:call-template name="TableHandler">
											<xsl:with-param name="parmVerTable" select="$parmVerTable"/>
											<xsl:with-param name ="custom" select="$custom"/>
											<xsl:with-param name ="mastersubtbl" select="$mastersubtbl"/>
											<xsl:with-param name ="characterStyle" select="$characterStyle"/>
										</xsl:call-template>
									</xsl:for-each>
									<!--Checking if not nested table then td is closed-->
									<xsl:if test="not(child::w:tbl) or not(count(child::w:tbl)=0)">
										<!--Closing the td tag-->
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/td','&gt;')"/>
									</xsl:if>
								</xsl:if>
							</xsl:for-each>
							<xsl:variable name="setRowspan" select="myObj:SetRowspan()"/>
							<!--Closing table row-->
						</tr>
					</xsl:if>
				</xsl:for-each>
				<!--Closing table body-->
			</tbody>
			<!--Closing Table-->
		</table>
	</xsl:template>

</xsl:stylesheet>