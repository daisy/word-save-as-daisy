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
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:myObj="urn:Daisy" exclude-result-prefixes="w pic wp dcterms xsi cp dc a r v dcmitype myObj o xsl dgm">
	<!--<xsl:param name="sOperators"/>
	<xsl:param name="sMinuses"/>
	<xsl:param name="sNumbers"/>
	<xsl:param name="sZeros"/>-->
	<xsl:output method="xml" indent="no" />
	<!--Storing the default language of the document from styles.xml-->
	<xsl:variable name="doclang" select="document('word/styles.xml')//w:styles/w:docDefaults/w:rPrDefault/w:rPr/w:lang/@w:val"/>
	<xsl:variable name="doclangbidi" select="document('word/styles.xml')//w:styles/w:docDefaults/w:rPrDefault/w:rPr/w:lang/@w:bidi"/>
	<xsl:variable name="doclangeastAsia" select="document('word/styles.xml')//w:styles/w:docDefaults/w:rPrDefault/w:rPr/w:lang/@w:eastAsia"/>
	<!--Template to create NoteReference for FootNote and EndNote
  It is taking two parameters varFootnote_Id and varNote_Class. varFootnote_Id 
  will contain the Reference id of either Footnote or Endnote.-->
	<xsl:template name="tmpProcessFootNote">
		<xsl:param name="varFootnote_Id"/>
		<xsl:param name="varNote_Class"/>
		<xsl:param name="characterStyle"/>
		<xsl:message terminate="no">progress:footnote</xsl:message>
		<!--Checking for matching reference Id for Fotnote and Endnote in footnote.xml
    or endnote.xml-->
		<xsl:if test="document('word/footnotes.xml')//w:footnotes/w:footnote[@w:id=$varFootnote_Id]or document('word/endnotes.xml')//w:endnotes/w:endnote[@w:id=$varFootnote_Id]">
			<noteref>
				<!--Creating the attribute idref for Noteref element and assining it a value.-->
				<xsl:attribute name="idref">
					<!--If Note_Class is Footnotereference then it will have footnote id value -->
					<xsl:if test="$varNote_Class='FootnoteReference'">
						<xsl:value-of select="concat('#footnote-',$varFootnote_Id - 1)"/>
					</xsl:if>
					<!--If Note_Class is Footnotereference then it will have footnote id value -->
					<xsl:if test="$varNote_Class='EndnoteReference'">
						<xsl:value-of select="concat('#endnote-',$varFootnote_Id - 1)"/>
					</xsl:if>
				</xsl:attribute>
				<!--Creating the attribute class for Noteref element and assinging it a value.-->
				<xsl:attribute name="class">
					<xsl:if test="$varNote_Class='FootnoteReference'">
						<xsl:value-of select="substring($varNote_Class,1,8)"/>
					</xsl:if>
					<!--Creating the attribute class for Noteref element and assinging it a value.-->
					<xsl:if test="$varNote_Class='EndnoteReference'">
						<xsl:value-of select="substring($varNote_Class,1,7)"/>
					</xsl:if>
				</xsl:attribute>
				<!--Checking for languages-->
				<xsl:if test="(w:rPr/w:lang) or (w:rPr/w:rFonts/@w:hint)">
					<xsl:attribute name="xml:lang">
						<xsl:call-template name="Languages">
							<xsl:with-param name="Attribute" select="'1'"/>
						</xsl:call-template>
					</xsl:attribute>
				</xsl:if>
				<xsl:value-of select="$varFootnote_Id - 1"/>
			</noteref>
		</xsl:if>
	</xsl:template>

	<!--Template to add EndNote-->
	<xsl:template name="tmpNote">
		<xsl:param name="endNoteId"/>
		<xsl:param name="vernote"/>
		<xsl:param name="characterStyle"/>
		<xsl:param name="sOperators"/>
		<xsl:param name="sMinuses"/>
		<xsl:param name="sNumbers"/>
		<xsl:param name="sZeros"/>
		<xsl:message terminate="no">progress:note</xsl:message>
		<!--Checking for EndNoteId greater than 1-->
		<xsl:if test="$endNoteId &gt; 1">
			<note>
				<!--Creating attribute ID for Note element-->
				<xsl:attribute name="id">
					<xsl:value-of select="concat('endnote-',$endNoteId - 1)"/>
				</xsl:attribute>
				<!--Creating attribute class for Note element-->
				<xsl:attribute name="class">
					<xsl:value-of select="'Endnote'"/>
				</xsl:attribute>
				<!--Travering each w:endnote element in endnote.xml file-->
				<xsl:for-each select="document('word/endnotes.xml')//w:endnotes/w:endnote">
					<xsl:message terminate="no">progress:parahandler</xsl:message>
					<!--Checks for matching Id-->
					<xsl:if test="@w:id=$endNoteId">
						<!--Travering each element inside w:endnote in endnote.xml file-->
						<xsl:for-each select="./node()">
							<xsl:message terminate="no">progress:parahandler</xsl:message>
							<!--Checking for Paragraph element-->
							<xsl:if test="name()='w:p'">
								<xsl:call-template name="ParagraphStyle">
									<xsl:with-param name="VERSION" select="$vernote"/>
									<xsl:with-param name="flagNote" select="'Note'"/>
									<xsl:with-param name="checkid" select="$endNoteId"/>
									<xsl:with-param name="sOperators" select="$sOperators"/>
									<xsl:with-param name="sMinuses" select="$sMinuses"/>
									<xsl:with-param name="sNumbers" select="$sNumbers"/>
									<xsl:with-param name="sZeros" select="$sZeros"/>
									<xsl:with-param name="characterparaStyle" select="$characterStyle"/>
								</xsl:call-template>
							</xsl:if>
						</xsl:for-each>
						<xsl:variable name="SetFlag" select="myObj:InitializeNoteFlag()"/>
					</xsl:if>
				</xsl:for-each>
			</note>
		</xsl:if>
	</xsl:template>

	<!--Template for Adding footnote-->
	<xsl:template name="footnote">
		<xsl:param name="verfoot"/>
		<xsl:param name="mastersubfoot"/>
		<xsl:param name="characterStyle"/>
		<xsl:param name="sOperators"/>
		<xsl:param name="sMinuses"/>
		<xsl:param name="sNumbers"/>
		<xsl:param name="sZeros"/>
		<xsl:message terminate="no">progressfootnote</xsl:message>
		<!--Inserting default footnote id in the array list-->
		<xsl:variable name="checkid" select="myObj:FootNoteId(0)"/>
		<!--Chaecking for the matching Id returned from footnoteId function of c#-->
		<xsl:if test="$checkid!=0">
			<!--Traversing through each footnote element in footnotes.xml file-->
			<xsl:for-each select="document('word/footnotes.xml')//w:footnotes/w:footnote">
				<xsl:message terminate="no">progress:parahandler</xsl:message>
				<!--Checking if Id returned from C# is equal to the footnote Id in footnotes.xml file-->
				<xsl:if test="@w:id=$checkid">
					<!--Creating note element and it's attribute values-->
					<note id="{concat('footnote-',$checkid - 1)}" class="Footnote">
						<!--Travering each element inside w:footnote in footnote.xml file-->
						<xsl:for-each select="./node()">
							<xsl:message terminate="no">progress:parahandler</xsl:message>
							<!--Checking for Paragraph element-->
							<xsl:if test="name()='w:p'">
								<xsl:choose>
									<!--Checking for MathImage in Word2003/xp  footnotes-->
									<xsl:when test="(w:r/w:object/v:shape/v:imagedata/@r:id) and (not(w:r/w:object/o:OLEObject[@ProgID='Equation.DSMT4']))" >
										<p>
											<xsl:value-of select="$checkid - 1"/>
											<imggroup>
												<img>
													<!--Variable to hold r:id from document.xml-->
													<xsl:variable name="Math_id">
														<xsl:value-of select="w:r/w:object/v:shape/v:imagedata/@r:id"/>
													</xsl:variable>
													<xsl:attribute name="alt">
														<xsl:choose>
															<!--Checking for alt text for MathEquation Image or providing
											  'Math Equation' as alttext-->
															<xsl:when test="w:r/w:object/v:shape/@alt">
																<xsl:value-of select="w:r/w:object/v:shape/@alt"/>
															</xsl:when>
															<xsl:otherwise>
																<xsl:value-of select ="'Math Equation'"/>
															</xsl:otherwise>
														</xsl:choose>
													</xsl:attribute>
													<!--Attribute holding the name of the Image-->
													<xsl:attribute name="src">
														<!--Caling MathImageFootnote for copying Image to output folder-->
														<xsl:value-of select ="myObj:MathImageFootnote($Math_id)"/>
													</xsl:attribute>
												</img>
											</imggroup>
										</p>
									</xsl:when>
									<xsl:when test="w:r/w:object/o:OLEObject[@ProgID='Equation.DSMT4']">
										<xsl:variable name="Math_DSMT4" select="myObj:GetMathML('wdFootnotesStory')"/>
										<xsl:choose>
											<xsl:when test="$Math_DSMT4=''">
												<imggroup>
													<img>
														<!--Creating variable mathimage for storing r:id value from document.xml-->
														<xsl:variable name="Math_rid">
															<xsl:value-of select="w:r/w:object/v:shape/v:imagedata/@r:id"/>
														</xsl:variable>
														<xsl:attribute name="alt">
															<xsl:choose>
																<!--Checking for alt Text-->
																<xsl:when test="w:r/w:object/v:shape/@alt">
																	<xsl:value-of select="w:r/w:object/v:shape/@alt"/>
																</xsl:when>
																<xsl:otherwise>
																	<!--Hardcoding value 'Math Equation'if user donot provide alt text for Math Equations-->
																	<xsl:value-of select ="'Math Equation'"/>
																</xsl:otherwise>
															</xsl:choose>
														</xsl:attribute>
														<xsl:attribute name="src">
															<!--Calling MathImage function-->
															<xsl:value-of select ="myObj:MathImageFootnote($Math_rid)"/>
														</xsl:attribute>
													</img>
												</imggroup>
											</xsl:when>
											<xsl:otherwise>
												<xsl:value-of disable-output-escaping="yes" select="$Math_DSMT4"/>
											</xsl:otherwise>
										</xsl:choose>
									</xsl:when>
									<xsl:otherwise>
										<!--Calling template for checking style in the footnote text-->
										<xsl:call-template name="ParagraphStyle">
											<xsl:with-param name="VERSION" select="$verfoot"/>
											<xsl:with-param name="flagNote" select="'Note'"/>
											<xsl:with-param name="checkid" select="$checkid"/>
											<xsl:with-param name="sOperators" select="$sOperators"/>
											<xsl:with-param name="sMinuses" select="$sMinuses"/>
											<xsl:with-param name="sNumbers" select="$sNumbers"/>
											<xsl:with-param name="sZeros" select="$sZeros"/>
											<xsl:with-param name="characterparaStyle" select="$characterStyle"/>
										</xsl:call-template>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:if>
						</xsl:for-each>
					</note>
				</xsl:if>
				<xsl:variable name="SetFlag" select="myObj:InitializeNoteFlag()"/>
			</xsl:for-each>
			<!--Calling the template footnote recursively until the C# function returns 0-->
			<xsl:call-template name="footnote">
				<xsl:with-param name ="verfoot" select ="$verfoot"/>
				<xsl:with-param name ="mastersubfoot" select ="$mastersubfoot"/>
				<xsl:with-param name="characStyle" select="$characterStyle"/>
				<xsl:with-param name="sOperators" select="$sOperators"/>
				<xsl:with-param name="sMinuses" select="$sMinuses"/>
				<xsl:with-param name="sNumbers" select="$sNumbers"/>
				<xsl:with-param name="sZeros" select="$sZeros"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<!--Template for handling multiple Prodnotes and Captions applied to an image-->
	<xsl:template name="ProcessCaptionProdNote">
		<xsl:param name="followingnodes"/>
		<xsl:param name="imageId"/>
		<xsl:param name="characterStyle"/>
		<xsl:choose>
			<!--Checking for inbuilt caption and Image-CaptionDAISY custom paragraph style-->
			<xsl:when test="($followingnodes[1]/w:pPr/w:pStyle/@w:val='Caption') or ($followingnodes[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
				<!--Variable holds the count of the captions and prodnotes-->
				<xsl:variable name="tmpcount" select="myObj:AddCaptionsProdnotes()"/>
				<caption>
					<!--attribute holds the value of the image id-->
					<xsl:attribute name="imgref">
						<xsl:value-of select="$imageId"/>
					</xsl:attribute>
					<xsl:if test="($followingnodes[1]/w:r/w:rPr/w:lang) or ($followingnodes[1]/w:r/w:rPr/w:rFonts/@w:hint)">
						<!--attribute holds the id of the language-->
						<xsl:attribute name="xml:lang">
							<xsl:call-template name="PictureLanguage">
								<xsl:with-param name="CheckLang" select="'picture'"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<!--Checking if image is bidirectionally oriented-->
					<xsl:if test="($followingnodes[1]/w:pPr/w:bidi) or ($followingnodes[1]/w:r/w:rPr/w:rtl)">
						<xsl:variable name="quote">"</xsl:variable>
						<!--Variable holds the value which indicates that the image is bidirectionally oriented-->
						<xsl:variable name="Bd">
							<!--calling the PictureLanguage template-->
							<xsl:call-template name="PictureLanguage">
								<xsl:with-param name="CheckLang" select="'picture'"/>
							</xsl:call-template>
						</xsl:variable>
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
					</xsl:if>
					<!--Looping through each of the node to print text to the output xml-->
					<xsl:for-each select="$followingnodes[1]/node()">
						<xsl:message terminate="no">progress:parahandler</xsl:message>
						<xsl:if test="name()='w:r'">
							<xsl:call-template name ="TempCharacterStyle">
								<xsl:with-param name ="characterStyle" select="$characterStyle"/>
							</xsl:call-template>
						</xsl:if>
						<xsl:if test="name()='w:fldSimple'">
							<xsl:value-of select="w:r/w:t"/>
						</xsl:if>

					</xsl:for-each>
					<!--Checking if image is bidirectionally oriented-->
					<xsl:if test="$followingnodes[1]/w:pPr/w:bidi">
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
					</xsl:if>
				</caption>
				<!--Recursively calling the ProcessCaptionProdNote template till all the Captions are processed-->
				<xsl:call-template name="ProcessCaptionProdNote">
					<xsl:with-param name="followingnodes" select="$followingnodes[position() > 1]"/>
					<xsl:with-param name="imageId" select="$imageId"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
			</xsl:when>
			<!--Checking for inbuilt caption and Prodnote-OptionalDAISY custom paragraph style-->
			<xsl:when test="($followingnodes[1]/w:pPr/w:pStyle/@w:val='Prodnote-OptionalDAISY')">
				<!--Variable holds the count of the captions and prodnotes-->
				<xsl:variable name="tmpcount" select="myObj:AddCaptionsProdnotes()"/>
				<xsl:variable name="quote">"</xsl:variable>
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','prodnote ','render= ',$quote,'optional',$quote,' imgref=',$quote,$imageId,$quote,'&gt;')"/>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="($followingnodes[1]/w:pPr/w:bidi) or ($followingnodes[1]/w:r/w:rPr/w:rtl)">
					<!--Variable holds the value which indicates that the image is bidirectionally oriented-->
					<xsl:variable name="Bd">
						<!--calling the PictureLanguage template-->
						<xsl:call-template name="PictureLanguage">
							<xsl:with-param name="CheckLang" select="'picture'"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
				</xsl:if>
				<!--Looping through each of the node to print text to the output xml-->
				<xsl:for-each select="$followingnodes[1]/node()">
					<xsl:message terminate="no">progress:parahandler</xsl:message>
					<xsl:if test="name()='w:r'">
						<xsl:call-template name ="TempCharacterStyle">
							<xsl:with-param name ="characterStyle" select="$characterStyle"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:for-each>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="$followingnodes[1]/w:pPr/w:bidi">
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
				</xsl:if>
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/prodnote ','&gt;')"/>
				<!--Recursively calling the ProcessCaptionProdNote template till all the ProdNotes are processed-->
				<xsl:call-template name="ProcessCaptionProdNote">
					<xsl:with-param name="followingnodes" select="$followingnodes[position() > 1]"/>
					<xsl:with-param name="imageId" select="$imageId"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
			</xsl:when>
			<!--Checking for inbuilt caption and Prodnote-RequiredDAISY custom paragraph style-->
			<xsl:when test="($followingnodes[1]/w:pPr/w:pStyle/@w:val='Prodnote-RequiredDAISY')">
				<!--Variable holds the count of the captions and prodnotes-->
				<xsl:variable name="tmpcount" select="myObj:AddCaptionsProdnotes()"/>
				<xsl:variable name="quote">"</xsl:variable>
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','prodnote ','render=',$quote,'required',$quote,' imgref=',$quote, $imageId ,$quote,'&gt;')"/>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="($followingnodes[1]/w:pPr/w:bidi) or ($followingnodes[1]/w:r/w:rPr/w:rtl)">
					<!--Variable holds the value which indicates that the image is bidirectionally oriented-->
					<xsl:variable name="Bd">
						<!--calling the PictureLanguage template-->
						<xsl:call-template name="PictureLanguage">
							<xsl:with-param name="CheckLang" select="'picture'"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
				</xsl:if>
				<!--Looping through each of the node to print text to the output xml-->
				<xsl:for-each select="$followingnodes[1]/node()">
					<xsl:message terminate="no">progress:parahandler</xsl:message>
					<xsl:if test="name()='w:r'">
						<xsl:call-template name ="TempCharacterStyle">
							<xsl:with-param name ="characterStyle" select="$characterStyle"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:for-each>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="$followingnodes[1]/w:pPr/w:bidi">
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
				</xsl:if>
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/prodnote ','&gt;')"/>
				<!--Recursively calling the ProcessCaptionProdNote template till all the ProdNotes are processed-->
				<xsl:call-template name="ProcessCaptionProdNote">
					<xsl:with-param name="followingnodes" select="$followingnodes[position() > 1]"/>
					<xsl:with-param name="imageId" select="$imageId"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
			</xsl:when>
		</xsl:choose>
		<xsl:variable name="tmpcount" select="myObj:ResetCaptionsProdnotes()"/>
	</xsl:template>

	<!--Template for implementing Simple Images i.e, ungrouped images-->
	<xsl:template name="PictureHandler">
		<xsl:param name="imgOpt"/>
		<xsl:param name="dpi"/>
		<xsl:param name="characterStyle"/>
		<xsl:message terminate="no">progress:picturehandler</xsl:message>
		<xsl:variable name="alttext">
			<xsl:value-of select="w:drawing/wp:inline/wp:docPr/@descr"/>
		</xsl:variable>
		<!--Variable holds the value of Image Id-->
		<xsl:variable name="Img_Id">
			<xsl:choose>
				<xsl:when  test="w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed">
					<xsl:value-of select="w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed">
					<xsl:value-of select="w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:inline/wp:docPr/@id">
					<xsl:value-of select="w:drawing/wp:inline/wp:docPr/@id"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/wp:docPr/@id">
					<xsl:value-of select="w:drawing/wp:anchor/wp:docPr/@id"/>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<!--Variable holds the filename of the image-->
		<xsl:variable name="imageName">
			<xsl:choose>
				<xsl:when  test="w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:nvPicPr/pic:cNvPr/@name">
					<xsl:value-of select="w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:nvPicPr/pic:cNvPr/@name"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic/pic:nvPicPr/pic:cNvPr/@name">
					<xsl:value-of select="w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic/pic:nvPicPr/pic:cNvPr/@name"/>
				</xsl:when>
			</xsl:choose>

		</xsl:variable>
		<!--Variable holds the value of Image Id concatenated with some random number generated for Image Id-->
		<xsl:variable name="imageId">
			<xsl:choose>
				<xsl:when  test="w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed">
					<xsl:value-of select="concat($Img_Id,myObj:GenerateImageId())"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed">
					<xsl:value-of select="concat($Img_Id,myObj:GenerateImageId())"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:inline/wp:docPr/@id">
					<xsl:variable name="id" select="../w:bookmarkStart[last()]/@w:name"/>
					<xsl:value-of select="myObj:CheckShapeId($id)"/>
				</xsl:when>
				<xsl:when test="contains(w:drawing/wp:inline/wp:docPr/@name,'Diagram')">
					<xsl:value-of select="myObj:CheckShapeId(concat('Shape',substring-after(../../../../@id,'s')))"/>
				</xsl:when>
				<xsl:when test="contains(w:drawing/wp:inline/wp:docPr/@name,'Chart')">
					<xsl:variable name="id" select="myObj:CheckShapeId(concat('Shape',../w:bookmarkStart[last()]/@w:name))"/>
				</xsl:when>
				<!--<xsl:when test="w:drawing/wp:anchor/wp:docPr/@id">
					<xsl:choose>
						<xsl:when test="contains(w:drawing/wp:anchor/wp:docPr/@name,'Chart')">
							<xsl:variable name="id" select="concat('Shape',w:drawing/wp:anchor/wp:docPr/@id)"/>
							<xsl:value-of select="myObj:CheckShapeId($id)"/>
						</xsl:when>
						<xsl:when test="contains(w:drawing/wp:anchor/wp:docPr/@name,'Diagram')">
							<xsl:value-of select="myObj:CheckShapeId(concat('Shape',w:drawing/wp:anchor/wp:docPr/@id))"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="myObj:CheckShapeId(concat('Shape',w:drawing/wp:anchor/wp:docPr/@id))"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>-->
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="imageWidth">
			<xsl:choose>
				<xsl:when  test="w:drawing/wp:inline/wp:extent">
					<xsl:value-of select="w:drawing/wp:inline/wp:extent/@cx"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/wp:extent">
					<xsl:value-of select="w:drawing/wp:anchor/wp:extent/@cx"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:inline/wp:extent">
					<xsl:value-of select="w:drawing/wp:inline/wp:extent/@cx"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/wp:extent">
					<xsl:value-of select="w:drawing/wp:anchor/wp:extent/@cx"/>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="imageHeight">
			<xsl:choose>
				<xsl:when  test="w:drawing/wp:inline/wp:extent">
					<xsl:value-of select="w:drawing/wp:inline/wp:extent/@cy"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/wp:extent">
					<xsl:value-of select="w:drawing/wp:anchor/wp:extent/@cy"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:inline/wp:extent">
					<xsl:value-of select="w:drawing/wp:inline/wp:extent/@cy"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/wp:extent">
					<xsl:value-of select="w:drawing/wp:anchor/wp:extent/@cy"/>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<!--Checking if Img_Id variable contains any Image Id-->
		<xsl:if test="string-length($Img_Id)>0">
			<!--Checking if document is bidirectionally oriented-->
			<xsl:if test="(../w:pPr/w:bidi) or (../w:pPr/w:jc/@w:val='right')">
				<xsl:variable name="quote">"</xsl:variable>
				<!--Variable holds the value which indicates that the image is bidirectionally oriented-->
				<xsl:variable name="imgBd">
					<!--calling the PictureLanguage template-->
					<xsl:call-template name="PictureLanguage">
						<xsl:with-param name="CheckLang" select="'picture'"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote, ' xml:lang=',$quote,$imgBd,$quote,'&gt;')"/>
			</xsl:if>
			<xsl:variable name="imageTest">
				<xsl:choose>
					<xsl:when test="contains($Img_Id,'rId') and ($imgOpt='resize')">
						<xsl:value-of select ="myObj:Image($Img_Id,$imageName,'true')"/>
					</xsl:when>
					<xsl:when test="contains($Img_Id,'rId') and ($imgOpt='resample')">
						<xsl:value-of select ="myObj:ResampleImage($Img_Id,$imageName,$dpi)"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<xsl:when test="contains($Img_Id,'rId')">
								<xsl:value-of select ="myObj:Image($Img_Id,$imageName,'true')"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="concat($imageId,'.png')"/>
							</xsl:otherwise>
						</xsl:choose>

					</xsl:otherwise>
				</xsl:choose>
			</xsl:variable>
			<xsl:variable name="checkImage">
				<xsl:value-of select="myObj:CheckImage($imageTest)"/>
			</xsl:variable>
			<xsl:if test="$checkImage='1'">
				<!--Creating Imagegroup element-->
				<imggroup>
					<img>
						<!--attribute that holds the value of the Image ID-->
						<xsl:attribute name="id">
							<xsl:value-of select="$imageId"/>
						</xsl:attribute>
						<!--attribute that holds the filename of the image returned for C# Image function-->
						<xsl:choose>
							<xsl:when test="$imgOpt='resize' and contains($Img_Id,'rId')">
								<xsl:attribute name="src">
									<xsl:value-of select ="$imageTest"/>
								</xsl:attribute>
								<!--attribute that holds the alternate text for the image-->
								<xsl:attribute name="alt">
									<xsl:value-of select="$alttext"/>
								</xsl:attribute>
								<xsl:attribute name="width">
									<xsl:value-of select="round(($imageWidth) div (9525))"/>
								</xsl:attribute>
								<xsl:attribute name="height">
									<xsl:value-of select="round(($imageHeight) div (9525))"/>
								</xsl:attribute>
							</xsl:when>
							<xsl:when test="$imgOpt='resample'  and contains($Img_Id,'rId')">
								<xsl:attribute name="src">
									<xsl:value-of select ="$imageTest"/>
								</xsl:attribute>
								<!--attribute that holds the alternate text for the image-->
								<xsl:attribute name="alt">
									<xsl:value-of select="$alttext"/>
								</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="src">
									<xsl:choose>
										<xsl:when test="contains($Img_Id,'rId')">
											<xsl:value-of select ="$imageTest"/>
										</xsl:when>
										<xsl:otherwise>
											<xsl:value-of select="$imageTest"/>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:attribute>
								<xsl:attribute name="alt">
									<xsl:value-of select="$alttext"/>
								</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
					</img>
					<!--Handling Image-CaptionDAISY custom paragraph style applied above an image-->
					<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY') or (../w:pPr/w:pStyle/@w:val='Caption') or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
						<caption>
							<xsl:attribute name="imgref">
								<xsl:value-of select="$imageId"/>
							</xsl:attribute>
							<xsl:if test="(../following-sibling::w:p[1]/w:r/w:rPr/w:lang) or (../following-sibling::w:p[1]/w:r/w:rPr/w:rFonts/@w:hint)">
								<xsl:attribute name="xml:lang">
									<xsl:call-template name="PictureLanguage">
										<xsl:with-param name="CheckLang" select="'picture'"/>
									</xsl:call-template>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="(../following-sibling::w:p[1]/w:pPr/w:bidi) or (../following-sibling::w:p[1]/w:r/w:rPr/w:rtl)">
								<xsl:variable name="quote">"</xsl:variable>
								<xsl:variable name="Bd">
									<xsl:call-template name="PictureLanguage">
										<xsl:with-param name="CheckLang" select="'picture'"/>
									</xsl:call-template>
								</xsl:variable>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p  ','xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
							</xsl:if>
							<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
								<xsl:for-each select="../preceding-sibling::node()[1]/node()">
									<xsl:message terminate="no">progress:parahandler</xsl:message>
									<!--Printing the Caption value-->
									<xsl:if test="name()='w:r'">
										<xsl:call-template name ="TempCharacterStyle">
											<xsl:with-param name ="characterStyle" select="$characterStyle"/>
										</xsl:call-template>
									</xsl:if>
									<xsl:if test="name()='w:fldSimple'">
										<xsl:value-of select="w:r/w:t"/>
									</xsl:if>

								</xsl:for-each>
								<xsl:text> </xsl:text>
							</xsl:if>
							<xsl:if test="(../w:pPr/w:pStyle/@w:val='Caption') or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
								<xsl:for-each select="../node()">
									<xsl:message terminate="no">progress:parahandler</xsl:message>
									<!--Printing the Caption value-->
									<xsl:if test="name()='w:r'">
										<xsl:call-template name ="TempCharacterStyle">
											<xsl:with-param name ="characterStyle" select="$characterStyle"/>
										</xsl:call-template>
									</xsl:if>
									<xsl:if test="name()='w:fldSimple'">
										<xsl:value-of select="w:r/w:t"/>
									</xsl:if>

								</xsl:for-each>
								<xsl:text> </xsl:text>
							</xsl:if>
							<xsl:if test="../following-sibling::w:p[1]/w:pPr/w:bidi">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
							</xsl:if>
						</caption>
					</xsl:if>
					<!--calling the template to handle multiple Prodnotes and Captions applied for an image-->
					<xsl:call-template name="ProcessCaptionProdNote">
						<xsl:with-param name="followingnodes" select="../following-sibling::node()"/>
						<xsl:with-param name="imageId" select="$imageId"/>
						<xsl:with-param name="characterStyle" select="$characterStyle"/>
					</xsl:call-template>
					<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Caption')">
						<xsl:message terminate="no">translation.oox2Daisy.ImageCaption</xsl:message>
					</xsl:if>
				</imggroup>
				<!--Checking if document is bidirectionally oriented-->
				<xsl:if test="(../w:pPr/w:bidi) or (../w:pPr/w:jc/@w:val='right')">
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
				</xsl:if>
			</xsl:if>
			<xsl:if test="$checkImage='0'">
				<xsl:message terminate="no">translation.oox2Daisy.Image</xsl:message>
			</xsl:if>
			<!--Checking if Img_Id contains null value and returns the fidelity loss message-->
		</xsl:if>
	</xsl:template>

	<!--Template for handling multiple Prodnotes and Captions applied for grouped images-->
	<xsl:template name="ProcessProdNoteImggroups">
		<xsl:param name="followingnodes"/>
		<xsl:param name="imageId"/>
		<xsl:param name="characterStyle"/>
		<xsl:choose>
			<!--Checking for Image-CaptionDAISY custom paragraph style-->
			<xsl:when test="($followingnodes[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
				<!--Variable that holds count of the captions and prodnotes-->
				<xsl:variable name="tmpcount" select="myObj:AddCaptionsProdnotes()"/>
				<caption>
					<!--attribute that holds image id returned from C# ReturnImageGroupId()-->
					<xsl:attribute name="imgref">
						<xsl:value-of select="$imageId"/>
					</xsl:attribute>
					<!--Getting the language id by calling the PictureLanguage template-->
					<xsl:if test="($followingnodes[1]/w:r/w:rPr/w:lang) or ($followingnodes[1]/w:r/w:rPr/w:rFonts/@w:hint)">
						<!--attribute that holds language id-->
						<xsl:attribute name="xml:lang">
							<!--calling the PictureLanguage template-->
							<xsl:call-template name="PictureLanguage">
								<xsl:with-param name="CheckLang" select="'imagegroup'"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<!--Checking if image is bidirectionally oriented-->
					<xsl:if test="($followingnodes[1]/w:pPr/w:bidi) or ($followingnodes[1]/w:r/w:rPr/w:rtl)">
						<xsl:variable name="quote">"</xsl:variable>
						<!--Variable holds the value which indicates that the image is bidirectionally oriented-->
						<xsl:variable name="Bd">
							<!--calling the PictureLanguage template-->
							<xsl:call-template name="PictureLanguage">
								<xsl:with-param name="CheckLang" select="'imagegroup'"/>
							</xsl:call-template>
						</xsl:variable>
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
					</xsl:if>
					<!--Looping through each of the node to print the text to the output xml-->
					<xsl:for-each select="$followingnodes[1]/node()">
						<xsl:message terminate="no">progress:parahandler</xsl:message>
						<xsl:if test="name()='w:r'">
							<xsl:call-template name ="TempCharacterStyle">
								<xsl:with-param name ="characterStyle" select="$characterStyle"/>
							</xsl:call-template>
						</xsl:if>
						<xsl:if test="name()='w:fldSimple'">
							<xsl:value-of select="w:r/w:t"/>
						</xsl:if>

					</xsl:for-each>
					<!--Checking for image is bidirectionally oriented-->
					<xsl:if test="$followingnodes[1]/w:pPr/w:bidi">
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
					</xsl:if>
				</caption>
				<!--Recursively calling the ProcessCaptionProdNote template till all the Captions are processed-->
				<xsl:call-template name="ProcessProdNoteImggroups">
					<xsl:with-param name="followingnodes" select="$followingnodes[position() > 1]"/>
					<xsl:with-param name="imageId" select="$imageId"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
			</xsl:when>
			<!--Checking for Prodnote-OptionalDAISY custom paragraph style-->
			<xsl:when test="($followingnodes[1]/w:pPr/w:pStyle/@w:val='Prodnote-OptionalDAISY')">
				<!--Variable that holds count of the captions and prodnotes-->
				<xsl:variable name="tmpcount" select="myObj:AddCaptionsProdnotes()"/>
				<xsl:variable name="quote">"</xsl:variable>
				<!--<xsl:variable name="imageId">
					<xsl:value-of select="myObj:ReturnImageGroupId()"/>
				</xsl:variable>-->
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','prodnote ','render=',$quote,'optional',$quote,' imgref=',$quote,$imageId,$quote,'&gt;')"/>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="($followingnodes[1]/w:pPr/w:bidi) or ($followingnodes[1]/w:r/w:rPr/w:rtl)">
					<!--Variable holds the value which indicates that the image is bidirectionally oriented-->
					<xsl:variable name="Bd">
						<!--Calling the PictureLanguage template-->
						<xsl:call-template name="PictureLanguage">
							<xsl:with-param name="CheckLang" select="'imagegroup'"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
				</xsl:if>
				<!--Looping through each of the node to print the text to the output xml-->
				<xsl:for-each select="$followingnodes[1]/node()">
					<xsl:message terminate="no">progress:parahandler</xsl:message>
					<xsl:if test="name()='w:r'">
						<xsl:call-template name ="TempCharacterStyle">
							<xsl:with-param name ="characterStyle" select="$characterStyle"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:for-each>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="$followingnodes[1]/w:pPr/w:bidi">
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
				</xsl:if>
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/prodnote ','&gt;')"/>
				<!--Recursively calling the ProcessCaptionProdNote template till all the prodnotes are processed-->
				<xsl:call-template name="ProcessProdNoteImggroups">
					<xsl:with-param name="followingnodes" select="$followingnodes[position() > 1]"/>
					<xsl:with-param name="imageId" select="$imageId"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
			</xsl:when>
			<!--Checking for Prodnote-RequiredDAISY custom paragraph style-->
			<xsl:when test="($followingnodes[1]/w:pPr/w:pStyle/@w:val='Prodnote-RequiredDAISY')">
				<!--Variable that holds count of the captions and prodnotes-->
				<xsl:variable name="tmpcount" select="myObj:AddCaptionsProdnotes()"/>
				<xsl:variable name="quote">"</xsl:variable>
				<!--<xsl:variable name="imageId">
					<xsl:value-of select="myObj:ReturnImageGroupId()"/>
				</xsl:variable>-->
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','prodnote ','render=',$quote,'required',$quote,' imgref=',$quote,$imageId,$quote,'&gt;')"/>
				<!--Getting the language id by calling the PictureLanguage template-->
				<xsl:if test="($followingnodes[1]/w:pPr/w:bidi) or ($followingnodes[1]/w:r/w:rPr/w:rtl)">
					<!--attribute that holds language id-->
					<xsl:variable name="Bd">
						<!--calling the PictureLanguage template-->
						<xsl:call-template name="PictureLanguage">
							<xsl:with-param name="CheckLang" select="'imagegroup'"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
				</xsl:if>
				<!--Looping through each of the node to print the text to the output xml-->
				<xsl:for-each select="$followingnodes[1]/node()">
					<xsl:message terminate="no">progress:parahandler</xsl:message>
					<xsl:if test="name()='w:r'">
						<xsl:call-template name ="TempCharacterStyle">
							<xsl:with-param name ="characterStyle" select="$characterStyle"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:for-each>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="$followingnodes[1]/w:pPr/w:bidi">
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
				</xsl:if>
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/prodnote ','&gt;')"/>
				<!--Recursively calling the ProcessCaptionProdNote template till all the prodnotes are processed-->
				<xsl:call-template name="ProcessProdNoteImggroups">
					<xsl:with-param name="followingnodes" select="$followingnodes[position() > 1]"/>
					<xsl:with-param name="imageId" select="$imageId"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
			</xsl:when>
		</xsl:choose>
		<xsl:variable name="tmpcount" select="myObj:ResetCaptionsProdnotes()"/>
	</xsl:template>

	<!--Template for Implementing grouped images-->
	<xsl:template name="Imagegroups">
		<xsl:param name="characterStyle"/>
		<xsl:message terminate="no">progress:imagegroups</xsl:message>
		<!--Handling Image-CaptionDAISY custom paragraph style applied above an image-->
		<xsl:if test="../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY'">
			<xsl:variable name="caption">
				<xsl:for-each select="../preceding-sibling::node()[1]/node()">
					<xsl:message terminate="no">progress:parahandler</xsl:message>
					<xsl:if test="name()='w:r'">
						<xsl:call-template name ="TempCharacterStyle">
							<xsl:with-param name ="characterStyle" select="$characterStyle"/>
						</xsl:call-template>
					</xsl:if>
					<xsl:if test="name()='w:fldSimple'">
						<xsl:value-of select="w:r/w:t"/>
					</xsl:if>

				</xsl:for-each>
			</xsl:variable>
			<xsl:variable name="InsertedCaption" select="myObj:InsertCaption($caption)"/>
		</xsl:if>
		<!--Looping through each pict element and storing the caption value in the caption variable-->

		<xsl:if test="../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:pPr/w:pStyle[@w:val='Caption']">
			<xsl:variable name="caption">
				<xsl:for-each select="../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/node()">
					<xsl:message terminate="no">progress:parahandler</xsl:message>
					<xsl:if test="name()='w:r'">
						<xsl:call-template name ="TempCharacterStyle">
							<xsl:with-param name ="characterStyle" select="$characterStyle"/>
						</xsl:call-template>
					</xsl:if>
					<xsl:if test="name()='w:fldSimple'">
						<xsl:value-of select="w:r/w:t"/>
					</xsl:if>

				</xsl:for-each>
			</xsl:variable>
			<!--Inserting the caption value in the Arraylist through insertcaption C# function-->
			<xsl:variable name="InsertedCaption" select="myObj:InsertCaption($caption)"/>
		</xsl:if>
		<xsl:variable name="Imageid">
			<xsl:value-of select ="myObj:CheckShapeId(concat('Shape',substring-after(w:pict/v:group/@id,'s')))"/>
		</xsl:variable>
		<xsl:variable name="checkImage">
			<xsl:value-of select="myObj:CheckImage(concat($Imageid,'.png'))"/>
		</xsl:variable>
		<xsl:if test="$checkImage='1'">
			<!--Checking for the presence of Images-->
			<imggroup>
				<img>
					<!--Creating attribute id of img element-->
					<xsl:attribute name="id">
						<xsl:value-of select="$Imageid"/>
					</xsl:attribute>
					<!--Creating attribute alt for alternate text of img element-->
					<xsl:attribute name="alt">
						<xsl:value-of select="w:pict/v:group/@alt"/>
					</xsl:attribute>
					<!--Creating attribute src of img element-->
					<xsl:attribute name="src">
						<xsl:value-of select ="concat($Imageid,'.png')"/>
					</xsl:attribute>
				</img>

				<!--Variable holds the caption value returned form C# function returncaption-->
				<xsl:variable name="checkcaption">
					<xsl:value-of select="myObj:ReturnCaption()"/>
				</xsl:variable>
				<!--Checking if checkcaption variables holds any value-->
				<xsl:if test="$checkcaption!='0'">
					<caption>
						<!--Creating imgref attribute and assinging it the value returned by C# returnImagegroupId -->
						<xsl:attribute name="imgref">
							<xsl:value-of select="$Imageid"/>
						</xsl:attribute>
						<!--Creating xml:lang and assinging it the value returned by PictureLanguage template-->
						<xsl:if test="(../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang) or ../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:rFonts/@w:hint">
							<xsl:attribute name="xml:lang">
								<xsl:call-template name="PictureLanguage">
									<xsl:with-param name="CheckLang" select="'imagegroup'"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
						<!--Checking if image is bidirectionally oriented-->
						<xsl:if test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:pPr/w:bidi or (../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:rtl)">
							<xsl:variable name="quote">"</xsl:variable>
							<xsl:variable name="Bd">
								<xsl:call-template name="PictureLanguage">
									<xsl:with-param name="CheckLang" select="'imagegroup'"/>
								</xsl:call-template>
							</xsl:variable>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
						</xsl:if>
						<xsl:value-of select="$checkcaption"/>
						<xsl:if test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:pPr/w:bidi">
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
						</xsl:if>
					</caption>
				</xsl:if>
				<!--calling the template to handle multiple Prodnotes and Captions applied for image groups-->
				<xsl:call-template name="ProcessProdNoteImggroups">
					<xsl:with-param name="followingnodes" select="../following-sibling::node()"/>
					<xsl:with-param name="imageId" select="$Imageid"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
			</imggroup>
		</xsl:if>
		<xsl:if test="$checkImage='0'">
			<xsl:message terminate="no">translation.oox2Daisy.Image</xsl:message>
		</xsl:if>
	</xsl:template>

	<xsl:template name="Imagegroup2003">
		<xsl:param name="characterStyle"/>
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<!--Variable that holds the Image Id-->
		<xsl:variable name="imageId">
			<xsl:value-of select="concat(w:pict/v:shape/v:imagedata/@r:id,myObj:GenerateImageId())"/>
		</xsl:variable>
		<!--Checking if image is bidirectionally oriented-->
		<xsl:if test="(../w:pPr/w:bidi) or (../w:pPr/w:jc/@w:val='right')">
			<xsl:variable name="quote">"</xsl:variable>
			<xsl:variable name="imgBd">
				<xsl:call-template name="PictureLanguage">
					<xsl:with-param name="CheckLang" select="'picture'"/>
				</xsl:call-template>
			</xsl:variable>
			<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote, ' xml:lang=',$quote,$imgBd,$quote,'&gt;')"/>
		</xsl:if>
		<xsl:variable name="checkImage">
			<xsl:value-of select="myObj:CheckImage(myObj:Image(w:pict/v:shape/v:imagedata/@r:id,w:pict/v:shape/v:imagedata/@o:title))"/>
		</xsl:variable>
		<xsl:if test="$checkImage='1'">
			<imggroup>
				<img>
					<!--attribute to store Image id-->
					<xsl:attribute name="id">
						<xsl:value-of select="$imageId"/>
					</xsl:attribute>
					<!--variable to store Image name-->
					<xsl:variable name="image2003Name">
						<xsl:value-of select="w:pict/v:shape/v:imagedata/@o:title"/>
					</xsl:variable>
					<!--variable to store Image id-->
					<xsl:variable name="rid">
						<xsl:value-of select="w:pict/v:shape/v:imagedata/@r:id"/>
					</xsl:variable>
					<!--Creating attribute src of img element-->
					<xsl:attribute name="src">
						<xsl:value-of select ="myObj:Image($rid,$image2003Name)"/>
					</xsl:attribute>
					<!--Creating attribute alt for alternate text of img element-->
					<xsl:attribute name="alt">
						<xsl:value-of select="w:pict/v:shape/@alt"/>
					</xsl:attribute>
				</img>
				<!--Handling Image-CaptionDAISY custom paragraph style applied above an image-->
				<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')or (../w:pPr/w:pStyle/@w:val='Caption') or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
					<caption>
						<xsl:attribute name="imgref">
							<xsl:value-of select="$imageId"/>
						</xsl:attribute>
						<xsl:if test="(../following-sibling::w:p[1]/w:r/w:rPr/w:lang) or (../following-sibling::w:p[1]/w:r/w:rPr/w:rFonts/@w:hint)">
							<xsl:attribute name="xml:lang">
								<xsl:call-template name="PictureLanguage">
									<xsl:with-param name="CheckLang" select="'picture'"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="(../following-sibling::w:p[1]/w:pPr/w:bidi) or (../following-sibling::w:p[1]/w:r/w:rPr/w:rtl)">
							<xsl:variable name="quote">"</xsl:variable>
							<xsl:variable name="Bd">
								<xsl:call-template name="PictureLanguage">
									<xsl:with-param name="CheckLang" select="'picture'"/>
								</xsl:call-template>
							</xsl:variable>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p  ','xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
						</xsl:if>
						<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
							<xsl:for-each select="../preceding-sibling::node()[1]/node()">
								<xsl:message terminate="no">progress:parahandler</xsl:message>
								<!--Printing the Caption value-->
								<xsl:if test="name()='w:r'">
									<xsl:call-template name ="TempCharacterStyle">
										<xsl:with-param name ="characterStyle" select="$characterStyle"/>
									</xsl:call-template>
								</xsl:if>
								<xsl:if test="name()='w:fldSimple'">
									<xsl:value-of select="w:r/w:t"/>
								</xsl:if>

							</xsl:for-each>
							<xsl:text> </xsl:text>
						</xsl:if>
						<xsl:if test="(../w:pPr/w:pStyle/@w:val='Caption') or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
							<xsl:for-each select="../node()">
								<xsl:message terminate="no">progress:parahandler</xsl:message>
								<!--Printing the Caption value-->
								<xsl:if test="name()='w:r'">
									<xsl:call-template name ="TempCharacterStyle">
										<xsl:with-param name ="characterStyle" select="$characterStyle"/>
									</xsl:call-template>
								</xsl:if>
								<xsl:if test="name()='w:fldSimple'">
									<xsl:value-of select="w:r/w:t"/>
								</xsl:if>

							</xsl:for-each>
							<xsl:text> </xsl:text>
						</xsl:if>
						<xsl:if test="../following-sibling::w:p[1]/w:pPr/w:bidi">
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
						</xsl:if>
					</caption>
				</xsl:if>
				<!--calling the template to handle multiple Prodnotes and Captions applied for an image-->
				<xsl:call-template name="ProcessCaptionProdNote">
					<xsl:with-param name="followingnodes" select="../following-sibling::node()"/>
					<xsl:with-param name="imageId" select="$imageId"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
				<!--Capturing Fidelity loss for Captions above the image-->
				<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Caption')">
					<xsl:message terminate="no">translation.oox2Daisy.ImageCaption</xsl:message>
				</xsl:if>
			</imggroup>
			<!--Checking if image is bidirectionally oriented-->
			<xsl:if test="(../w:pPr/w:bidi) or (../w:pPr/w:jc/@w:val='right')">
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
			</xsl:if>
		</xsl:if>
		<xsl:if test="$checkImage='0'">
			<xsl:message terminate="no">translation.oox2Daisy.Image</xsl:message>
		</xsl:if>
	</xsl:template>


	<xsl:template name="Object">
		<xsl:param name="characterStyle"/>
		<xsl:if test="not(contains(w:object/o:OLEObject/@ProgID,'Equation'))">
			<xsl:variable name="quote">"</xsl:variable>
			<xsl:if test="(contains(w:object/o:OLEObject/@ProgID,'Excel')) or (contains(w:object/o:OLEObject/@ProgID,'Word')) or (contains(w:object/o:OLEObject/@ProgID,'PowerPoint'))">
				<xsl:variable name="href">
					<xsl:value-of select="myObj:Object(w:object/o:OLEObject/@r:id)"/>
				</xsl:variable>
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','a ','href=',$quote,$href,$quote,' ','external=',$quote,'true',$quote,'&gt;')"/>
			</xsl:if>
			<xsl:variable name="ImageName" select="myObj:MathImage(w:object/v:shape/v:imagedata/@r:id)"/>
			<xsl:variable name ="id" select="myObj:GenerateObjectId()"/>
			<xsl:variable name="ImageId" select="concat($ImageName,$id)"/>
			<xsl:variable name="checkImage" select="myObj:CheckImage($ImageName)"/>
			<xsl:if test="$checkImage='1'">
				<imggroup>
					<img>
						<xsl:attribute name="id">
							<xsl:value-of select="$ImageId"/>
						</xsl:attribute>
						<xsl:attribute name="alt">
							<xsl:choose>
								<xsl:when test="string-length(w:object/v:shape/@alt)!=0">
									<xsl:value-of select="w:object/v:shape/@alt"/>
								</xsl:when>
								<xsl:otherwise>
									<xsl:value-of select="w:object/o:OLEObject/@ProgID"/>
								</xsl:otherwise>
							</xsl:choose>

						</xsl:attribute>
						<xsl:attribute name="src">
							<xsl:value-of select="$ImageName"/>
						</xsl:attribute>
					</img>
					<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY') or (../w:pPr/w:pStyle/@w:val='Caption') or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
						<caption>
							<xsl:attribute name="imgref">
								<xsl:value-of select="$ImageId"/>
							</xsl:attribute>
							<xsl:if test="(../following-sibling::w:p[1]/w:r/w:rPr/w:lang) or (../following-sibling::w:p[1]/w:r/w:rPr/w:rFonts/@w:hint)">
								<xsl:attribute name="xml:lang">
									<xsl:call-template name="PictureLanguage">
										<xsl:with-param name="CheckLang" select="'picture'"/>
									</xsl:call-template>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="(../following-sibling::w:p[1]/w:pPr/w:bidi) or (../following-sibling::w:p[1]/w:r/w:rPr/w:rtl)">
								<xsl:variable name="Bd">
									<xsl:call-template name="PictureLanguage">
										<xsl:with-param name="CheckLang" select="'picture'"/>
									</xsl:call-template>
								</xsl:variable>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
							</xsl:if>
							<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
								<xsl:for-each select="../preceding-sibling::node()[1]/node()">
									<xsl:message terminate="no">progress:parahandler</xsl:message>
									<!--Printing the Caption value-->
									<xsl:if test="name()='w:r'">
										<xsl:call-template name ="TempCharacterStyle">
											<xsl:with-param name ="characterStyle" select="$characterStyle"/>
										</xsl:call-template>
									</xsl:if>
									<xsl:if test="name()='w:fldSimple'">
										<xsl:value-of select="w:r/w:t"/>
									</xsl:if>

								</xsl:for-each>
							</xsl:if>
							<xsl:if test="(../w:pPr/w:pStyle/@w:val='Caption') or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
								<xsl:for-each select="../node()">
									<xsl:message terminate="no">progress:parahandler</xsl:message>
									<!--Printing the Caption value-->
									<xsl:if test="name()='w:r'">
										<xsl:call-template name ="TempCharacterStyle">
											<xsl:with-param name ="characterStyle" select="$characterStyle"/>
										</xsl:call-template>
									</xsl:if>
									<xsl:if test="name()='w:fldSimple'">
										<xsl:value-of select="w:r/w:t"/>
									</xsl:if>

								</xsl:for-each>
								<xsl:text> </xsl:text>
							</xsl:if>
							<xsl:if test="../following-sibling::w:p[1]/w:pPr/w:bidi">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
							</xsl:if>
							<!--Printing the field value of the Caption-->
						</caption>
					</xsl:if>
					<xsl:call-template name="ProcessCaptionProdNote">
						<xsl:with-param name="followingnodes" select="../following-sibling::node()"/>
						<xsl:with-param name="imageId" select="$ImageId"/>
						<xsl:with-param name="characterStyle" select="$characterStyle"/>
					</xsl:call-template>

				</imggroup>
				<xsl:if test="(../w:pPr/w:bidi) or (../w:pPr/w:jc/@w:val='right')">
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
				</xsl:if>
				<xsl:if test="contains(w:object/o:OLEObject/@ProgID,'Excel') or contains(w:object/o:OLEObject/@ProgID,'Word') or contains(w:object/o:OLEObject/@ProgID,'PowerPoint')">
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/a','&gt;')"/>
				</xsl:if>
			</xsl:if>
			<xsl:if test="$checkImage='0'">
				<xsl:message terminate="no">translation.oox2Daisy.Image</xsl:message>
			</xsl:if>
		</xsl:if>
	</xsl:template>

	<xsl:template name="tmpShape">
		<xsl:param name="characterStyle"/>
		<xsl:variable name="imageId">
			<xsl:choose>
				<xsl:when test="(w:pict/v:shape/@id) and (w:pict/v:shape/@o:spid)">
					<xsl:value-of select="myObj:CheckShapeId(concat('Shape',substring-after(w:pict/v:shape/@o:spid,'s')))"/>
				</xsl:when>
				<xsl:when test="w:pict/v:shape/@id">
					<xsl:value-of select="myObj:CheckShapeId(concat('Shape',substring-after(w:pict/v:shape/@id,'s')))"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="myObj:CheckShapeId(concat('Shape',substring-after(w:pict//@id,'s')))"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="checkImage">
			<xsl:value-of select="myObj:CheckImage(concat($imageId,'.png'))"/>
		</xsl:variable>
		<xsl:if test="$checkImage='1'">
			<imggroup>
				<img>
					<xsl:attribute name="id">
						<xsl:value-of select="$imageId"/>
					</xsl:attribute>
					<xsl:attribute name="src">
						<xsl:value-of select="concat($imageId,'.png')"/>
					</xsl:attribute>
					<xsl:attribute name="alt">
						<xsl:value-of select="w:pict/v:shape/@alt"/>
					</xsl:attribute>
				</img>

				<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY') or (../w:pPr/w:pStyle/@w:val='Caption') or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
					<caption>
						<xsl:attribute name="imgref">
							<xsl:value-of select="$imageId"/>
						</xsl:attribute>
						<xsl:if test="(../following-sibling::w:p[1]/w:r/w:rPr/w:lang) or (../following-sibling::w:p[1]/w:r/w:rPr/w:rFonts/@w:hint)">
							<xsl:attribute name="xml:lang">
								<xsl:call-template name="PictureLanguage">
									<xsl:with-param name="CheckLang" select="'picture'"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="(../following-sibling::w:p[1]/w:pPr/w:bidi) or (../following-sibling::w:p[1]/w:r/w:rPr/w:rtl)">
							<xsl:variable name="quote">"</xsl:variable>
							<xsl:variable name="Bd">
								<xsl:call-template name="PictureLanguage">
									<xsl:with-param name="CheckLang" select="'picture'"/>
								</xsl:call-template>
							</xsl:variable>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p  ','xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','bdo ','dir= ',$quote,'rtl',$quote,' xml:lang=',$quote,$Bd,$quote,'&gt;')"/>
						</xsl:if>
						<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
							<xsl:for-each select="../preceding-sibling::node()[1]/node()">
								<xsl:message terminate="no">progress:parahandler</xsl:message>

								<!--Printing the Caption value-->

								<xsl:if test="name()='w:r'">
									<xsl:call-template name ="TempCharacterStyle">
										<xsl:with-param name ="characterStyle" select="$characterStyle"/>
									</xsl:call-template>
								</xsl:if>
								<xsl:if test="name()='w:fldSimple'">
									<xsl:value-of select="w:r/w:t"/>
								</xsl:if>

							</xsl:for-each>
							<xsl:text> </xsl:text>
						</xsl:if>
						<xsl:if test="(../w:pPr/w:pStyle/@w:val='Caption') or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
							<xsl:for-each select="../node()">
								<xsl:message terminate="no">progress:parahandler</xsl:message>

								<!--Printing the Caption value-->


								<xsl:if test="name()='w:r'">
									<xsl:call-template name ="TempCharacterStyle">
										<xsl:with-param name ="characterStyle" select="$characterStyle"/>
									</xsl:call-template>
								</xsl:if>
								<xsl:if test="name()='w:fldSimple'">
									<xsl:value-of select="w:r/w:t"/>
								</xsl:if>

							</xsl:for-each>
							<xsl:text> </xsl:text>
						</xsl:if>
						<xsl:if test="../following-sibling::w:p[1]/w:pPr/w:bidi">
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
						</xsl:if>
					</caption>
					<xsl:if test="../following-sibling::w:p[1]/w:pPr/w:bidi">
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/p','&gt;')"/>
					</xsl:if>
				</xsl:if>
				<xsl:call-template name="ProcessCaptionProdNote">
					<xsl:with-param name="followingnodes" select="../following-sibling::node()"/>
					<xsl:with-param name="imageId" select="$imageId"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
			</imggroup>
			<xsl:if test="(../w:pPr/w:bidi) or (../w:pPr/w:jc/@w:val='right')">
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','/bdo','&gt;')"/>
			</xsl:if>
		</xsl:if>
		<xsl:if test="$checkImage='0'">
			<xsl:message terminate="no">translation.oox2Daisy.Image</xsl:message>
		</xsl:if>
	</xsl:template>

	<!--Template for checking section breaks for page numbers-->
	<xsl:template name="SectionBreak">
		<xsl:param name="count"/>
		<xsl:param name="node"/>
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:variable name="initialize" select="myObj:InitalizeCheckSectionBody()"/>
		<xsl:variable name="reSetConPageBreak" select="myObj:ResetSetConPageBreak()"/>
		<xsl:choose>
			<!--if page number for front matter-->
			<xsl:when test="$node='front'">
				<!--incrementing the default page counter-->
				<xsl:variable name="increment" select="myObj:IncrementPage()"/>
				<!--Traversing through each node-->
				<xsl:for-each select="following-sibling::node()">
					<xsl:message terminate="no">progress:parahandler</xsl:message>
					<xsl:choose>
						<!--Checking for paragraph section break-->
						<xsl:when test="w:pPr/w:sectPr">

							<xsl:if test="myObj:CheckSectionBody()=1">
								<xsl:choose>
									<!--Checking if page start and page format is present-->
									<xsl:when test="(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="$count"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page format is present and not page start-->
									<xsl:when test="(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and not(w:pPr/w:sectPr/w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="$count"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page start is present and not page format-->
									<xsl:when test="not(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="$count"/>
										</xsl:call-template>
									</xsl:when>
									<!--If both are not present-->
									<xsl:otherwise>
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="$count"/>
										</xsl:call-template>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:if>
						</xsl:when>
						<!--Checking for Section in a document-->
						<xsl:when test="name()='w:sectPr'">
							<xsl:if test="myObj:CheckSectionBody()=1">
								<xsl:variable name="setSectionFront" select="myObj:CheckSectionFront()"/>
								<!--<xsl:variable name="frontCounter" select="myObj:IncrementPageNo()"/>-->
								<xsl:choose>
									<!--Checking if page start and page format is present-->
									<xsl:when test="(w:pgNumType/@w:fmt) and (w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="$count"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page format is present and not page start-->
									<xsl:when test="(w:pgNumType/@w:fmt) and not(w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="$count"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page start is present and not page format-->
									<xsl:when test="not(w:pgNumType/@w:fmt) and (w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="$count"/>
										</xsl:call-template>
									</xsl:when>
									<!--If both are not present-->
									<xsl:otherwise>
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="$count"/>
										</xsl:call-template>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:if>
						</xsl:when>
					</xsl:choose>
				</xsl:for-each>
			</xsl:when>
		</xsl:choose>
		<xsl:choose>
			<!--if page number for body matter-->
			<xsl:when test="$node='body'">
				<xsl:if test="../preceding-sibling::node()[1]/w:pPr/w:sectPr">
					<xsl:variable name="setConPageBreak" select="myObj:SetConPageBreak()"/>
				</xsl:if>
				<!--Traversing through each node-->
				<xsl:for-each select="../following-sibling::node()">
					<xsl:message terminate="no">progress:parahandler</xsl:message>
					<xsl:choose>
						<!--Checking for paragraph section break-->
						<xsl:when test="w:pPr/w:sectPr">
							<xsl:if test="myObj:CheckSectionBody()=1">
								<xsl:choose>
									<!--Checking if page start and page format is present-->
									<xsl:when test="(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page format is present and not page start-->
									<xsl:when test="(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and not(w:pPr/w:sectPr/w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="'0'"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page start is present and not page format-->
									<xsl:when test="not(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--If both are not present-->
									<xsl:otherwise>
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="'0'"/>
										</xsl:call-template>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:if>
						</xsl:when>
						<!--Checking for Section in a document-->
						<xsl:when test="name()='w:sectPr'">
							<xsl:if test="myObj:CheckSectionBody()=1">
								<xsl:choose>
									<!--Checking if page start and page format is present-->
									<xsl:when test="(w:pgNumType/@w:fmt) and (w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page format is present and not page start-->
									<xsl:when test="(w:pgNumType/@w:fmt) and not(w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="'0'"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page start is present and not page format-->
									<xsl:when test="not(w:pgNumType/@w:fmt) and (w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--If both are not present-->
									<xsl:otherwise>
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="'0'"/>
										</xsl:call-template>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:if>
						</xsl:when>
					</xsl:choose>
				</xsl:for-each>
			</xsl:when>
		</xsl:choose>
		<xsl:choose>
			<!--Checking for paragraph-->
			<xsl:when test="$node='Para'">
				<xsl:if test="preceding-sibling::node()[1]/w:pPr/w:sectPr">
					<xsl:variable name="setConPageBreak" select="myObj:SetConPageBreak()"/>
				</xsl:if>
				<!--Traversing through each node-->
				<xsl:for-each select="following-sibling::node()">
					<xsl:message terminate="no">progress:parahandler</xsl:message>
					<xsl:choose>
						<!--Checking for paragraph section break-->
						<xsl:when test="w:pPr/w:sectPr">
							<xsl:if test="myObj:CheckSectionBody()=1">
								<xsl:choose>
									<!--Checking if page start and page format is present-->
									<xsl:when test="(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page format is present and not page start-->
									<xsl:when test="(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and not(w:pPr/w:sectPr/w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="'0'"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page start is present and not page format-->
									<xsl:when test="not(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--If both are not present-->
									<xsl:otherwise>
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="'0'"/>
										</xsl:call-template>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:if>
						</xsl:when>
						<!--Checking for Section in a document-->
						<xsl:when test="name()='w:sectPr'">
							<xsl:if test="myObj:CheckSectionBody()=1">
								<xsl:choose>
									<!--Checking if page start and page format is present-->
									<xsl:when test="(w:pgNumType/@w:fmt) and (w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page format is present and not page start-->
									<xsl:when test="(w:pgNumType/@w:fmt) and not(w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="'0'"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page start is present and not page format-->
									<xsl:when test="not(w:pgNumType/@w:fmt) and (w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--If both are not present-->
									<xsl:otherwise>
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="'0'"/>
										</xsl:call-template>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:if>
						</xsl:when>
					</xsl:choose>
				</xsl:for-each>
			</xsl:when>
		</xsl:choose>
		<xsl:choose>
			<!--Checking for bodysection-->
			<xsl:when test="$node='bodysection'">
				<xsl:if test="preceding-sibling::node()[1]/w:pPr/w:sectPr">
					<xsl:variable name="setConPageBreak" select="myObj:SetConPageBreak()"/>
				</xsl:if>
				<xsl:choose>
					<!--Checking for paragraph section break-->
					<xsl:when test="w:pPr/w:sectPr">
						<xsl:choose>
							<!--Checking if page start and page format is present-->
							<xsl:when test="(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
								<!--Calling template for page number text-->
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
									<xsl:with-param name="matter" select="$node"/>
									<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
								</xsl:call-template>
							</xsl:when>
							<!--Checking if page format is present and not page start-->
							<xsl:when test="(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and not(w:pPr/w:sectPr/w:pgNumType/@w:start)">
								<!--Calling template for page number text-->
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
									<xsl:with-param name="matter" select="$node"/>
									<xsl:with-param name="counter" select="'0'"/>
								</xsl:call-template>
							</xsl:when>
							<!--Checking if page start is present and not page format-->
							<xsl:when test="not(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
								<!--Calling template for page number text-->
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
									<xsl:with-param name="matter" select="$node"/>
									<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
								</xsl:call-template>
							</xsl:when>
							<!--If both are not present-->
							<xsl:otherwise>
								<!--Calling template for page number text-->
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
									<xsl:with-param name="matter" select="$node"/>
									<xsl:with-param name="counter" select="'0'"/>
								</xsl:call-template>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
				</xsl:choose>
			</xsl:when>
		</xsl:choose>
		<xsl:choose>
			<!--Checking for Table-->
			<xsl:when test="$node='Table'">
				<xsl:if test="../../preceding-sibling::node()[1]/w:pPr/w:sectPr">
					<xsl:variable name="setConPageBreak" select="myObj:SetConPageBreak()"/>
				</xsl:if>
				<!--Traversing through each node-->
				<xsl:for-each select="../../following-sibling::node()">
					<xsl:message terminate="no">progress:parahandler</xsl:message>
					<xsl:choose>
						<!--Checking for paragraph section break-->
						<xsl:when test="w:pPr/w:sectPr">
							<xsl:choose>
								<!--Checking if page start and page format is present-->
								<xsl:when test="(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
									<!--Calling template for page number text-->
									<xsl:call-template name="PageNumber">
										<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
										<xsl:with-param name="matter" select="$node"/>
										<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
									</xsl:call-template>
								</xsl:when>
								<!--Checking if page format is present and not page start-->
								<xsl:when test="(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and not(w:pPr/w:sectPr/w:pgNumType/@w:start)">
									<!--Calling template for page number text-->
									<xsl:call-template name="PageNumber">
										<xsl:with-param name="pagetype" select="w:pPr/w:sectPr/w:pgNumType/@w:fmt"/>
										<xsl:with-param name="matter" select="$node"/>
										<xsl:with-param name="counter" select="'0'"/>
									</xsl:call-template>
								</xsl:when>
								<!--Checking if page start is present and not page format-->
								<xsl:when test="not(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
									<!--Calling template for page number text-->
									<xsl:call-template name="PageNumber">
										<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
										<xsl:with-param name="matter" select="$node"/>
										<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
									</xsl:call-template>
								</xsl:when>
								<!--If both are not present-->
								<xsl:otherwise>
									<!--Calling template for page number text-->
									<xsl:call-template name="PageNumber">
										<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
										<xsl:with-param name="matter" select="$node"/>
										<xsl:with-param name="counter" select="'0'"/>
									</xsl:call-template>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<!--Checking for Section in a document-->
						<xsl:when test="name()='w:sectPr'">
							<xsl:if test="myObj:CheckSectionBody()=1">
								<xsl:choose>
									<!--Checking if page start and page format is present-->
									<xsl:when test="(w:pgNumType/@w:fmt) and (w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="'body'"/>
											<xsl:with-param name="counter" select="w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page format is present and not page start-->
									<xsl:when test="(w:pgNumType/@w:fmt) and not(w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="w:pgNumType/@w:fmt"/>
											<xsl:with-param name="matter" select="'body'"/>
											<xsl:with-param name="counter" select="'0'"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page start is present and not page format-->
									<xsl:when test="not(w:pgNumType/@w:fmt) and (w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="myObj:GetPageFormat()"/>
											<xsl:with-param name="matter" select="'body'"/>
											<xsl:with-param name="counter" select="w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--If both are not present-->
									<xsl:otherwise>
										<!--Calling template for page number text-->
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
			</xsl:when>
		</xsl:choose>
	</xsl:template>

	<!--Template for counting number of pages before TOC-->
	<xsl:template name="countpageTOC">
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:for-each select="preceding-sibling::*">
			<xsl:message terminate="no">progress:parahandler</xsl:message>
			<xsl:choose>
				<!--Checking for page break in TOC-->
				<xsl:when test="(w:r/w:br/@w:type='page') or (w:r/w:lastRenderedPageBreak)">
					<xsl:variable name="check" select="myObj:PageForTOC()"/>
					<xsl:variable name="increment" select="myObj:IncrementPage()"/>
					<xsl:if test="not(w:r/w:t)">
						<!--Calling template for initializing page number info-->
						<xsl:call-template name="SectionBreak">
							<xsl:with-param name="count" select="myObj:ReturnPageNum()"/>
							<xsl:with-param name="node" select="'front'"/>
						</xsl:call-template>
						<!--producer note for empty text-->
						<prodnote>
							<xsl:attribute name="render">optional</xsl:attribute>
							<xsl:value-of select="'Blank Page'"/>
						</prodnote>
					</xsl:if>
				</xsl:when>
				<xsl:when test="(w:sdtContent/w:p/w:r/w:br/@w:type='page') or (w:sdtContent/w:p/w:r/lastRenderedPageBreak)">
					<xsl:variable name="check" select="myObj:PageForTOC()"/>
					<xsl:variable name="increment" select="myObj:IncrementPage()"/>
				</xsl:when>
			</xsl:choose>
		</xsl:for-each>
		<xsl:variable name="countPage" select="myObj:PageForTOC() - 1"/>
		<xsl:call-template name="SectionBreak">
			<xsl:with-param name="count" select="$countPage"/>
			<xsl:with-param name="node" select="'front'"/>
		</xsl:call-template>
	</xsl:template>

	<!--Template to translate page number information-->
	<xsl:template name="PageNumber">
		<xsl:param name="pagetype"/>
		<xsl:param name="matter"/>
		<xsl:param name="counter"/>
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:choose>
			<xsl:when test="myObj:GetCurrentMatterType()='Frontmatter'">
				<xsl:if test="not((myObj:SetConPageBreak()&gt;1) and (w:type/@w:val='continuous'))">
					<xsl:variable name="count" select="myObj:IncrementPageNo()-1"/>
					<xsl:choose>
						<!--LowerRoman page number-->
						<xsl:when test="$pagetype='lowerRoman'">
							<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="myObj:PageNumLowerRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperRoman page number-->
						<xsl:when test="$pagetype='upperRoman'">
							<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="myObj:PageNumUpperRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--LowerLetter page number-->
						<xsl:when test="$pagetype='lowerLetter'">
							<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="myObj:PageNumLowerAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperLetter page number-->
						<xsl:when test="$pagetype='upperLetter'">
							<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="myObj:PageNumUpperAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--Page number with dash-->
						<xsl:when test="$pagetype='numberInDash'">
							<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="concat('-',$count,'-')"/>
							</pagenum>
						</xsl:when>
						<!--Normal page number-->
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test="$counter='0' and myObj:GetSectionFront()=1">
									<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
										<xsl:value-of select="$count"/>
									</pagenum>
								</xsl:when>
								<xsl:when test="$counter='0' and myObj:GetSectionFront()=0">
									<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
										<xsl:value-of select="myObj:ReturnPageNum()"/>
									</pagenum>
								</xsl:when>
								<xsl:otherwise>
									<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
										<xsl:value-of select="$count"/>
									</pagenum>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>
			</xsl:when>
      
			<xsl:when test="myObj:GetCurrentMatterType()='Bodymatter'">
				<xsl:if test="not((myObj:SetConPageBreak()&gt;1) and (w:type/@w:val='continuous'))">
					<xsl:variable name="count" select="myObj:IncrementPageNo()-1"/>
          <xsl:choose>
            <!--LowerRoman page number-->
            <xsl:when test="$pagetype='lowerRoman'">
              <pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
                <xsl:value-of select="myObj:PageNumLowerRoman($count)"/>
              </pagenum>
            </xsl:when>
            <!--UpperRoman page number-->
            <xsl:when test="$pagetype='upperRoman'">
              <pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
                <xsl:value-of select="myObj:PageNumUpperRoman($count)"/>
              </pagenum>
            </xsl:when>
            <!--LowerLetter page number-->
            <xsl:when test="$pagetype='lowerLetter'">
              <pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
                <xsl:value-of select="myObj:PageNumLowerAlphabet($count)"/>
              </pagenum>
            </xsl:when>
            <!--UpperLetter page number-->
            <xsl:when test="$pagetype='upperLetter'">
              <pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
                <xsl:value-of select="myObj:PageNumUpperAlphabet($count)"/>
              </pagenum>
            </xsl:when>
            <!--Page number with dash-->
            <xsl:when test="$pagetype='numberInDash'">
              <pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
                <xsl:value-of select="concat('-',$count,'-')"/>
              </pagenum>
            </xsl:when>
            <!--Normal page number-->
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="$counter='0' and myObj:GetSectionFront()=1">
                  <pagenum page="normal" id="{concat('page',myObj:GeneratePageId())}">
                    <xsl:value-of select="$count"/>
                  </pagenum>
                </xsl:when>
                <xsl:when test="$counter='0' and myObj:GetSectionFront()=0">
                  <pagenum page="normal" id="{concat('page',myObj:GeneratePageId())}">
                    <xsl:value-of select="myObj:ReturnPageNum()"/>
                  </pagenum>
                </xsl:when>
                <xsl:otherwise>
                  <pagenum page="normal" id="{concat('page',myObj:GeneratePageId())}">
                    <xsl:value-of select="$count"/>
                  </pagenum>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
				</xsl:if>
			</xsl:when>

			<xsl:when test="myObj:GetCurrentMatterType()='Reartmatter'">
				<xsl:if test="not((myObj:SetConPageBreak()&gt;1) and (w:type/@w:val='continuous'))">
					<xsl:variable name="count" select="myObj:IncrementPageNo()-1"/>
					<xsl:choose>
						<!--LowerRoman page number-->
						<xsl:when test="$pagetype='lowerRoman'">
							<pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="myObj:PageNumLowerRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperRoman page number-->
						<xsl:when test="$pagetype='upperRoman'">
							<pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="myObj:PageNumUpperRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--LowerLetter page number-->
						<xsl:when test="$pagetype='lowerLetter'">
							<pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="myObj:PageNumLowerAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperLetter page number-->
						<xsl:when test="$pagetype='upperLetter'">
							<pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="myObj:PageNumUpperAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--Page number with dash-->
						<xsl:when test="$pagetype='numberInDash'">
							<pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="concat('-',$count,'-')"/>
							</pagenum>
						</xsl:when>
						<!--Normal page number-->
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test="$counter='0' and myObj:GetSectionFront()=1">
									<pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
										<xsl:value-of select="$count"/>
									</pagenum>
								</xsl:when>
								<xsl:when test="$counter='0' and myObj:GetSectionFront()=0">
									<pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
										<xsl:value-of select="myObj:ReturnPageNum()"/>
									</pagenum>
								</xsl:when>
								<xsl:otherwise>
									<pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
										<xsl:value-of select="$count"/>
									</pagenum>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>
			</xsl:when>
      
			<!--Frontmatter page number-->
			<xsl:when test="$matter='front'">
				<xsl:if test="myObj:GetSectionFront()=1">
					<xsl:variable name="count" select="myObj:IncrementPageNo()"/>
				</xsl:if>
				<xsl:choose>
					<!--LowerRoman page number-->
					<xsl:when test="$pagetype='lowerRoman'">
						<xsl:variable name="pageno" select="myObj:PageNumLowerRoman($counter)"/>
						<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
							<xsl:value-of select="$pageno"/>
						</pagenum>
					</xsl:when>
					<!--UpperRoman page number-->
					<xsl:when test="$pagetype='upperRoman'">
						<xsl:variable name="pageno" select="myObj:PageNumUpperRoman($counter)"/>
						<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
							<xsl:value-of select="$pageno"/>
						</pagenum>
					</xsl:when>
					<!--LowerLetter page number-->
					<xsl:when test="$pagetype='lowerLetter'">
						<xsl:variable name="pageno" select="myObj:PageNumLowerAlphabet($counter)"/>
						<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
							<xsl:value-of select="$pageno"/>
						</pagenum>
					</xsl:when>
					<!--UpperLetter page number-->
					<xsl:when test="$pagetype='upperLetter'">
						<xsl:variable name="pageno" select="myObj:PageNumUpperAlphabet($counter)"/>
						<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
							<xsl:value-of select="$pageno"/>
						</pagenum>
					</xsl:when>
					<!--Page number with dash-->
					<xsl:when test="$pagetype='numberInDash'">
						<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
							<xsl:value-of select="concat('-',$counter,'-')"/>
						</pagenum>
					</xsl:when>
					<!--Normal page number-->
					<xsl:otherwise>
						<pagenum page="front" id="{concat('page',myObj:GeneratePageId())}">
							<xsl:value-of select="$counter"/>
						</pagenum>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
      
			<!--Bodymatter page number-->
			<xsl:when test="($matter='body') or ($matter='bodysection') or ($matter='Para')">
				<xsl:if test="not((myObj:SetConPageBreak()&gt;1) and (w:type/@w:val='continuous'))">
					<xsl:variable name="count" select="myObj:IncrementPageNo()-1"/>
					<xsl:choose>
						<!--LowerRoman page number-->
						<xsl:when test="$pagetype='lowerRoman'">
							<pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="myObj:PageNumLowerRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperRoman page number-->
						<xsl:when test="$pagetype='upperRoman'">
							<pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="myObj:PageNumUpperRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--LowerLetter page number-->
						<xsl:when test="$pagetype='lowerLetter'">
							<pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="myObj:PageNumLowerAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperLetter page number-->
						<xsl:when test="$pagetype='upperLetter'">
							<pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="myObj:PageNumUpperAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--Page number with dash-->
						<xsl:when test="$pagetype='numberInDash'">
							<pagenum page="special" id="{concat('page',myObj:GeneratePageId())}">
								<xsl:value-of select="concat('-',$count,'-')"/>
							</pagenum>
						</xsl:when>
						<!--Normal page number-->
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test="$counter='0' and myObj:GetSectionFront()=1">
									<pagenum page="normal" id="{concat('page',myObj:GeneratePageId())}">
										<xsl:value-of select="$count"/>
									</pagenum>
								</xsl:when>
								<xsl:when test="$counter='0' and myObj:GetSectionFront()=0">
									<pagenum page="normal" id="{concat('page',myObj:GeneratePageId())}">
										<xsl:value-of select="myObj:ReturnPageNum()"/>
									</pagenum>
								</xsl:when>
								<xsl:otherwise>
									<pagenum page="normal" id="{concat('page',myObj:GeneratePageId())}">
										<xsl:value-of select="$count"/>
									</pagenum>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>
			</xsl:when>
		</xsl:choose>
	</xsl:template>

	<!--Template to implement Languages-->
	<xsl:template name="Languages">
		<xsl:param name="Attribute"/>
		<xsl:param name ="txt"/>
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:variable name="quote">"</xsl:variable>
		<xsl:variable name="count_lang">      
      <xsl:for-each select="w:r[1]/w:rPr/w:lang">
				<xsl:value-of select="count(@*)"/>
			</xsl:for-each>
		</xsl:variable>
		<xsl:choose>
			<!--Checking for language type CS-->
			<xsl:when test="w:r/w:rPr/w:rFonts/@w:hint='cs'">
				<xsl:choose>
					<!--Checking for bidirectional language-->
					<xsl:when test="w:r/w:rPr/w:lang/@w:bidi">
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->
							<xsl:when test="$Attribute='0'">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:bidi,$quote,'&gt;')"/>
							</xsl:when>
							<xsl:otherwise>
								<!--Assingning language value-->
								<xsl:value-of select="w:r/w:rPr/w:lang/@w:bidi"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->
							<xsl:when test="$Attribute='0'">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$doclangbidi,$quote,'&gt;')"/>
							</xsl:when>
							<xsl:otherwise>
								<!--Assingning default language value-->
								<xsl:value-of select="$doclangbidi"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<!--Checking for language type eastAsia-->
			<xsl:when test="w:r/w:rPr/w:rFonts/@w:hint='eastAsia'">
				<xsl:choose>
					<xsl:when test="w:r/w:rPr/w:lang/@w:eastAsia">
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->
							<xsl:when test="$Attribute='0'">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:eastAsia,$quote,'&gt;')"/>
							</xsl:when>
							<xsl:otherwise>
								<!--Assingning language value-->
								<xsl:value-of select="w:r/w:rPr/w:lang/@w:eastAsia"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->
							<xsl:when test="$Attribute='0'">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$doclangeastAsia,$quote,'&gt;')"/>
							</xsl:when>
							<xsl:otherwise>
								<!--Assingning default language value-->
								<xsl:value-of select="$doclangeastAsia"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="$count_lang&gt;1">
						<xsl:choose>
							<xsl:when test="w:r/w:rPr/w:lang/@w:val">
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->
									<xsl:when test="$Attribute='0'">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:val,$quote,'&gt;')"/>
									</xsl:when>
									<!--Assingning language value-->
									<xsl:otherwise>
										<xsl:value-of select="w:r/w:rPr/w:lang/@w:val"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->
									<xsl:when test="$Attribute='0'">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$doclang,$quote,'&gt;')"/>
									</xsl:when>
									<!--Assingning default language value-->
									<xsl:otherwise>
										<xsl:value-of select="$doclang"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="$count_lang=1">
						<xsl:choose>
							<xsl:when test="w:r/w:rPr/w:lang/@w:val">
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->
									<xsl:when test="$Attribute='0'">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:val,$quote,'&gt;')"/>
									</xsl:when>
									<!--Assingning language value-->
									<xsl:otherwise>
										<xsl:value-of select="w:r/w:rPr/w:lang/@w:val"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="w:r/w:rPr/w:lang/@w:eastAsia">
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->
									<xsl:when test="$Attribute='0'">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:eastAsia,$quote,'&gt;')"/>
									</xsl:when>
									<!--Assingning language value-->
									<xsl:otherwise>
										<xsl:value-of select="w:r/w:rPr/w:lang/@w:eastAsia"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="w:r/w:rPr/w:lang/@w:bidi">
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->
									<xsl:when test="$Attribute='0'">

										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:bidi,$quote,'&gt;')"/>
									</xsl:when>
									<!--Assingning language value-->
									<xsl:otherwise>
										<xsl:value-of select="w:r/w:rPr/w:lang/@w:bidi"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->
							<xsl:when test="$Attribute='0'">
								<xsl:choose>
									<xsl:when test="w:r/w:rPr/w:lang/@w:val">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:val,$quote,'&gt;')"/>
									</xsl:when>
									<xsl:when test="w:r/w:rPr/w:lang/@w:eastAsia">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:eastAsia,$quote,'&gt;')"/>
									</xsl:when>
									<xsl:when test="w:r/w:rPr/w:lang/@w:bidi">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:bidi,$quote,'&gt;')"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p','&gt;')"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<!--Assingning default language value-->
							<xsl:otherwise>
								<xsl:value-of select="$doclang"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!--Template to implement Languages-->
	<xsl:template name="LanguagesPara">
		<xsl:param name="Attribute"/>
		<xsl:param name ="level"/>
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:variable name="quote">"</xsl:variable>
		<xsl:variable name="count_lang">
      <!--NOTE: Use w:r instead w:r[1]-->
			<xsl:for-each select="w:r/w:rPr/w:lang">
				<xsl:value-of select="count(@*)"/>
			</xsl:for-each>
		</xsl:variable>
		<xsl:choose>
			<!--Checking for language type CS-->

			<xsl:when test="w:r/w:rPr/w:rFonts/@w:hint='cs'">
				<xsl:choose>
					<!--for bidirectional language-->
					<xsl:when test="w:r/w:rPr/w:lang/@w:bidi">
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->
							<xsl:when test="$Attribute='0'">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:bidi,$quote,'&gt;')"/>
							</xsl:when>
							<!--Creating <level> element with xml:lang attribute-->
							<xsl:otherwise>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=',$quote,w:r/w:rPr/w:lang/@w:bidi,$quote,'&gt;')"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->

							<xsl:when test="$Attribute='0'">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$doclangbidi,$quote,'&gt;')"/>
							</xsl:when>
							<!--Creating <level> element with xml:lang attribute-->

							<xsl:otherwise>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=',$quote,$doclangbidi,$quote,'&gt;')"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<!--Checking for language type eastAsia-->

			<xsl:when test="w:r/w:rPr/w:rFonts/@w:hint='eastAsia'">
				<xsl:choose>
					<xsl:when test="w:r/w:rPr/w:lang/@w:eastAsia">
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->
							<xsl:when test="$Attribute='0'">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:eastAsia,$quote,'&gt;')"/>
							</xsl:when>
							<!--Creating <level> element with xml:lang attribute-->

							<xsl:otherwise>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=',$quote,w:r/w:rPr/w:lang/@w:eastAsia,$quote,'&gt;')"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->

							<xsl:when test="$Attribute='0'">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$doclangeastAsia,$quote,'&gt;')"/>
							</xsl:when>
							<!--Creating <level> element with xml:lang attribute-->

							<xsl:otherwise>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=',$quote,$doclangeastAsia,$quote,'&gt;')"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="$count_lang&gt;1">
						<xsl:choose>
							<xsl:when test="w:r/w:rPr/w:lang/@w:val">
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->

									<xsl:when test="$Attribute='0'">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:val,$quote,'&gt;')"/>
									</xsl:when>
									<!--Assingning language value-->

									<xsl:otherwise>
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=',$quote,w:r/w:rPr/w:lang/@w:val,$quote,'&gt;')"/>
									</xsl:otherwise>
									<!--<xsl:otherwise>
										<xsl:value-of select="w:r/w:rPr/w:lang/@w:val"/>
									</xsl:otherwise>-->
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->

									<xsl:when test="$Attribute='0'">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,$doclang,$quote,'&gt;')"/>
									</xsl:when>
									<!--Assingning default language value-->

									<xsl:otherwise>
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=',$quote,$doclang,$quote,'&gt;')"/>
										<!--<xsl:value-of select="$doclang"/>-->
									</xsl:otherwise>
								</xsl:choose>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="$count_lang=1">
						<xsl:choose>
							<xsl:when test="w:r/w:rPr/w:lang/@w:val">
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->
									<xsl:when test="$Attribute='0'">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:val,$quote,'&gt;')"/>
									</xsl:when>
									<!--Assingning language value-->

									<xsl:otherwise>
										<!--<xsl:value-of select="w:r/w:rPr/w:lang/@w:val"/>-->
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=',$quote,w:r/w:rPr/w:lang/@w:val,$quote,'&gt;')"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="w:r/w:rPr/w:lang/@w:eastAsia">
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->
									<xsl:when test="$Attribute='0'">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:eastAsia,$quote,'&gt;')"/>
									</xsl:when>
									<!--Assingning language value-->
									<xsl:otherwise>
										<!--<xsl:value-of select="w:r/w:rPr/w:lang/@w:eastAsia"/>-->
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=',$quote,w:r/w:rPr/w:lang/@w:eastAsia,$quote,'&gt;')"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="w:r/w:rPr/w:lang/@w:bidi">
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->
									<xsl:when test="$Attribute='0'">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p ','xml:lang=',$quote,w:r/w:rPr/w:lang/@w:bidi,$quote,'&gt;')"/>
									</xsl:when>
									<!--Assingning language value-->
									<xsl:otherwise>
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=',$quote,w:r/w:rPr/w:lang/@w:bidi,$quote,'&gt;')"/>
										<!--<xsl:value-of select="w:r/w:rPr/w:lang/@w:bidi"/>-->
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->
									<xsl:when test="$Attribute='0'">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p','&gt;')"/>
									</xsl:when>
									<!--Assingning default language value-->
									<xsl:otherwise>
										<xsl:value-of select="$doclang"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->
							<xsl:when test="$Attribute='0'">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;','p','&gt;')"/>
							</xsl:when>
							<!--Assingning default language value-->
							<xsl:otherwise>
								<xsl:value-of select="$doclang"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!--Template to implement Languages-->
	<xsl:template name="PictureLanguage">
		<xsl:param name="CheckLang"/>
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:choose>
			<!--Checking languge for picture-->
			<xsl:when test="$CheckLang='picture'">
				<xsl:variable name="count_lang">
					<xsl:for-each select="../following-sibling::w:p[1]/w:r[1]/w:rPr/w:lang">
						<xsl:value-of select="count(@*)"/>
					</xsl:for-each>
				</xsl:variable>
				<xsl:choose>
					<!--Checking for language type eastAsia-->
					<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:rFonts/@w:hint='eastAsia'">
						<xsl:choose>
							<!--Getting value from eastasia attribute in lang tag-->
							<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia">
								<xsl:value-of select="(../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia)"/>
							</xsl:when>
							<!--Assinging default eastAsia language-->
							<xsl:otherwise>
								<xsl:value-of select="$doclangeastAsia"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<!--Checking for language type CS-->
					<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:rFonts/@w:hint='cs'">
						<xsl:choose>
							<!--Checking for bidirectional language-->
							<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi">
								<!--Getting value from bidi attribute in lang tag-->
								<xsl:value-of select="(../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi)"/>
							</xsl:when>
							<!--Assinging default bidirectional language-->
							<xsl:otherwise>
								<xsl:value-of select="$doclangbidi"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<xsl:when test="$count_lang &gt;1">
								<xsl:choose>
									<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val">
										<xsl:value-of select="(../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val)"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="$doclang"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="$count_lang=1">
								<xsl:choose>
									<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val">
										<xsl:value-of select="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val"/>
									</xsl:when>
									<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia">
										<xsl:value-of select="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia"/>
									</xsl:when>
									<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi">
										<xsl:value-of select="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi"/>
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$doclang"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<!--Checking language for image group-->
			<xsl:when test="$CheckLang='imagegroup'">
				<xsl:variable name="count_lang">
					<xsl:for-each select="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang">
						<xsl:value-of select="count(@*)"/>
					</xsl:for-each>
				</xsl:variable>
				<xsl:choose>
					<!--Checking for language type CS-->
					<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:rFonts/@w:hint='cs'">
						<xsl:choose>
							<!--Checking for bidirectional language-->
							<xsl:when test="(../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:bidi)">
								<!--Getting value from bidi attribute in lang tag-->
								<xsl:value-of select="(../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:bidi)"/>
							</xsl:when>
							<!--Assinging default bidirectional language-->
							<xsl:otherwise>
								<xsl:value-of select="$doclangbidi"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<!--Checking for language type eastAsia-->
					<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:rFonts/@w:hint='eastAsia'">
						<xsl:choose>
							<!--Getting value from eastasia attribute in lang tag-->
							<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:eastAsia">
								<xsl:value-of select="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:eastAsia"/>
							</xsl:when>
							<!--Assinging default eastAsia language-->
							<xsl:otherwise>
								<xsl:value-of select="$doclangeastAsia"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<xsl:when test="$count_lang &gt; 1">
								<xsl:choose>
									<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:val">
										<xsl:value-of select="(../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:val)"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="$doclang"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="$count_lang=1">
								<xsl:choose>
									<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:val">
										<xsl:value-of select="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:val"/>
									</xsl:when>
									<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:eastAsia">
										<xsl:value-of select="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:eastAsia"/>
									</xsl:when>
									<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:bidi">
										<xsl:value-of select="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:bidi"/>
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$doclang"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<!--Checking language for table-->
			<xsl:when test="$CheckLang='Table'">
				<xsl:variable name="count_lang">
					<xsl:for-each select="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang">
						<xsl:value-of select="count(@*)"/>
					</xsl:for-each>
				</xsl:variable>
				<xsl:choose>
					<!--Checking for language type eastAsia-->
					<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:rFonts/@w:hint='eastAsia'">
						<xsl:choose>
							<!--Getting value from eastasia attribute in lang tag-->

							<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia">
								<xsl:value-of select="(preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia)"/>
							</xsl:when>
							<!--Assinging default eastAsia language-->

							<xsl:otherwise>
								<xsl:value-of select="$doclangeastAsia"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<!--Checking for language type CS-->
					<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:rFonts/@w:hint='cs'">
						<xsl:choose>
							<!--Checking for bidirectional language-->
							<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi">
								<!--Getting value from bidi attribute in lang tag-->
								<xsl:value-of select="(preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi)"/>
							</xsl:when>
							<!--Assinging default bidirectional language-->
							<xsl:otherwise>
								<xsl:value-of select="$doclangbidi"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<xsl:when test="$count_lang &gt; 1">
								<xsl:choose>
									<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val">
										<xsl:value-of select="(preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val)"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="$doclang"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="$count_lang = 1">
								<xsl:choose>
									<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val">
										<xsl:value-of select="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val"/>
									</xsl:when>
									<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia">
										<xsl:value-of select="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia"/>
									</xsl:when>
									<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi">
										<xsl:value-of select="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi"/>
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$doclang"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
		</xsl:choose>
	</xsl:template>

	<!--Template for taking language for bdo Tag-->
	<xsl:template name="BdoLanguages">
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:variable name="quote">"</xsl:variable>
		<xsl:variable name="count_lang">
			<xsl:for-each select="../../w:r[1]/w:rPr/w:lang">
				<xsl:value-of select="count(@*)"/>
			</xsl:for-each>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="../../w:r/w:rPr/w:rFonts/@w:hint='cs'">
				<xsl:choose>
					<xsl:when test="../../w:r/w:rPr/w:lang/@w:bidi">
						<xsl:value-of select="../../w:r/w:rPr/w:lang/@w:bidi"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="$doclangbidi"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:when test="../../w:r/w:rPr/w:rFonts/@w:hint='eastAsia'">
				<xsl:choose>
					<xsl:when test="../../w:r/w:rPr/w:lang/@w:eastAsia">
						<xsl:value-of select="../../w:r/w:rPr/w:lang/@w:eastAsia"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="$doclangeastAsia"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="$count_lang&gt;1">
						<xsl:choose>
							<xsl:when test="../../w:r/w:rPr/w:lang/@w:val">
								<xsl:value-of select="../../w:r/w:rPr/w:lang/@w:val"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$doclang"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="$count_lang=1">
						<xsl:choose>
							<xsl:when test="../../w:r/w:rPr/w:lang/@w:val">
								<xsl:value-of select="../../w:r/w:rPr/w:lang/@w:val"/>
							</xsl:when>
							<xsl:when test="../../w:r/w:rPr/w:lang/@w:bidi">
								<xsl:value-of select="../../w:r/w:rPr/w:lang/@w:bidi"/>
							</xsl:when>
							<xsl:when test="../../w:r/w:rPr/w:lang/@w:eastAsia">
								<xsl:value-of select="../../w:r/w:rPr/w:lang/@w:eastAsia"/>
							</xsl:when>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="$doclang"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--Template for taking language for bdo Tag-->
	<xsl:template name="BdoRtlLanguages">
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:variable name="quote">"</xsl:variable>
		<xsl:variable name="count_lang">
			<xsl:for-each select="../w:r[1]/w:rPr/w:lang">
				<xsl:value-of select="count(@*)"/>
			</xsl:for-each>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="w:rPr/w:rFonts/@w:hint='cs'">
				<xsl:choose>
					<xsl:when test="w:rPr/w:lang/@w:bidi">
						<xsl:value-of select="w:rPr/w:lang/@w:bidi"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="$doclangbidi"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:when test="w:rPr/w:rFonts/@w:hint='eastAsia'">
				<xsl:choose>
					<xsl:when test="w:rPr/w:lang/@w:eastAsia">
						<xsl:value-of select="w:rPr/w:lang/@w:eastAsia"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="$doclangeastAsia"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="$count_lang&gt;1">
						<xsl:choose>
							<xsl:when test="w:rPr/w:lang/@w:val">
								<xsl:value-of select="w:rPr/w:lang/@w:val"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$doclang"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="$count_lang=1">
						<xsl:choose>
							<xsl:when test="w:rPr/w:lang/@w:val">
								<xsl:value-of select="w:rPr/w:lang/@w:val"/>
							</xsl:when>
							<xsl:when test="w:rPr/w:lang/@w:bidi">
								<xsl:value-of select="w:rPr/w:lang/@w:bidi"/>
							</xsl:when>
							<xsl:when test="w:rPr/w:lang/@w:eastAsia">
								<xsl:value-of select="w:rPr/w:lang/@w:eastAsia"/>
							</xsl:when>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="$doclang"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="TempCharacterStyle">
		<xsl:param name ="characterStyle"/>
		<xsl:message terminate="no">progress:parahandler</xsl:message>
		<xsl:choose>
			<xsl:when test="$characterStyle='True'">
				<xsl:choose>
					<xsl:when test="../w:pPr/w:ind[@w:left] and ../w:pPr/w:ind[@w:right] and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:caps and w:rPr/w:color and w:t">
						<xsl:variable name="val" select="../w:pPr/w:ind/@w:left"/>
						<xsl:variable name="val_left" select="($val div 1440)"/>
						<xsl:variable name="valright" select="../w:pPr/w:ind/@w:right"/>
						<xsl:variable name="val_right" select="($valright div 1440)"/>
						<xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';text-transform:uppercase',';text-indent:','right=',$val_right,'in',';left=',$val_left,'in')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="../w:pPr/w:ind[@w:left] and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:caps and w:rPr/w:color and w:t">
						<xsl:variable name="val" select="../w:pPr/w:ind/@w:left"/>
						<xsl:variable name="val_left" select="($val div 1440)"/>
						<xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';text-transform:uppercase',';text-indent:',$val_left,'in')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="../w:pPr/w:ind[@w:right] and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:caps and w:rPr/w:color and w:t">
						<xsl:variable name="val" select="../w:pPr/w:ind/@w:right"/>
						<xsl:variable name="val_right" select="($val div 1440)"/>
						<xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';text-transform:uppercase',';text-indent:',$val_right,'in')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="../w:pPr/w:jc and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:caps and w:rPr/w:color and w:t">
						<xsl:variable name="val" select="../w:pPr/w:jc/@w:val"/>
						<xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';text-transform:uppercase',';text-align:',$val)}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="w:rPr/w:u and w:rPr/w:strike and w:rPr/w:caps and w:rPr/w:color and w:t">
						<xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';text-transform:uppercase')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>


					<xsl:when test="../w:pPr/w:ind[@w:left] and ../w:pPr/w:ind[@w:right] and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:smallCaps and w:rPr/w:color and w:t">
						<xsl:variable name="val" select="../w:pPr/w:ind/@w:left"/>
						<xsl:variable name="val_left" select="($val div 1440)"/>
						<xsl:variable name="valright" select="../w:pPr/w:ind/@w:right"/>
						<xsl:variable name="val_right" select="($valright div 1440)"/>
						<xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';font-variant:small-caps',';text-indent:','right=',$val_right,'in',';left=',$val_left,'in')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="../w:pPr/w:ind[@w:left] and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:smallCaps and w:rPr/w:color and w:t">
						<xsl:variable name="val" select="../w:pPr/w:ind/@w:left"/>
						<xsl:variable name="val_left" select="($val div 1440)"/>
						<xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';font-variant:small-caps',';text-indent:',$val_left,'in')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="../w:pPr/w:ind[@w:right] and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:smallCaps and w:rPr/w:color and w:t">
						<xsl:variable name="val" select="../w:pPr/w:ind/@w:right"/>
						<xsl:variable name="val_right" select="($val div 1440)"/>
						<xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';font-variant:small-caps',';text-indent:',$val_right,'in')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="../w:pPr/w:jc and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:smallCaps and w:rPr/w:color and w:t">
						<xsl:variable name="val" select="../w:pPr/w:jc/@w:val"/>
						<xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';font-variant:small-caps',';text-align:',$val)}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="w:rPr/w:u and w:rPr/w:strike and w:rPr/w:smallCaps and w:rPr/w:color and w:t">
						<xsl:variable name="val_color" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';font-variant:small-caps')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="w:rPr/w:u and w:t">
						<span class="text-decoration: underline">
							<xsl:value-of disable-output-escaping="yes" select="w:t"/>
						</span>
					</xsl:when>
					<xsl:when test="w:rPr/w:strike and w:t">
						<span class="text-decoration:line-through">
							<xsl:value-of disable-output-escaping="yes" select="concat(' ',w:t)"/>
						</span>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="w:t"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="w:t"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

</xsl:stylesheet>