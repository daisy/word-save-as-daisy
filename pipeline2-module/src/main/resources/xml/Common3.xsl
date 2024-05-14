<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="2.0"
                xmlns:xs="http://www.w3.org/2001/XMLSchema"
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
                xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                xmlns:o="urn:schemas-microsoft-com:office:office"
                xmlns:d="org.daisy.pipeline.word_to_dtbook.impl.DaisyClass"
                xmlns="http://www.daisy.org/z3986/2005/dtbook/"
                exclude-result-prefixes="w pic wp dcterms xsi cp dc a r v dcmitype d o xsl dgm">
	<!--Storing the default language of the document from styles.xml-->
	<xsl:variable name="doclang" as="xs:string" select="$stylesXml//w:styles/w:docDefaults/w:rPrDefault/w:rPr/w:lang/@w:val"/>
	<xsl:variable name="doclangbidi" as="xs:string" select="$stylesXml//w:styles/w:docDefaults/w:rPrDefault/w:rPr/w:lang/@w:bidi"/>
	<xsl:variable name="doclangeastAsia" as="xs:string" select="$stylesXml//w:styles/w:docDefaults/w:rPrDefault/w:rPr/w:lang/@w:eastAsia"/>
	<xsl:param name="Title"/>
	<!--Holds Documents Title value-->
	<xsl:param name="Creator"/>
	<!--Holds Documents creator value-->
	<xsl:param name="Publisher"/>
	<!--Holds Documents Publisher value-->
	<xsl:param name="UID"/>
	<!--Holds Document unique id value-->
	<xsl:param name="Subject"/>
	<!--Holds Documents Subject value-->
	<xsl:param name="acceptRevisions"/>
	<xsl:param name="Version"/>
	<!--Holds Documents version value-->
	<xsl:param name="Custom"/>
	<xsl:param name="MasterSub"/>
	<xsl:param name="ImageSizeOption"/>
	<xsl:param name="DPI"/>
	<xsl:param name="CharacterStyles"/>
	<xsl:param name="FootnotesPosition"/>
	<xsl:param name="FootnotesLevel"/>
	<xsl:param name="FootnotesNumbering" />
	<xsl:param name="FootnotesStartValue" />
	<xsl:param name="FootnotesNumberingPrefix" />
	<xsl:param name="FootnotesNumberingSuffix" />
	<xsl:param name="Language" />
	<!--Template to create NoteReference for FootNote and EndNote
  It is taking two parameters varFootnote_Id and varNote_Class. varFootnote_Id 
  will contain the Reference id of either Footnote or Endnote.-->
	<xsl:template name="NoteReference">
		<xsl:param name="noteID" as="xs:integer"/>
		<xsl:param name="noteClass" as="xs:string"/>
		<xsl:message terminate="no">progress:footnote</xsl:message>
		<!--Checking for matching reference Id for Fotnote and Endnote in footnote.xml
	or endnote.xml-->
		<xsl:if test="$footnotesXml//w:footnotes/w:footnote[@w:id=$noteID]or $endnotesXml//w:endnotes/w:endnote[@w:id=$noteID]">
			<noteref>
				<!--Creating the attribute idref for Noteref element and assining it a value.-->
				<xsl:attribute name="idref">
					<!--If Note_Class is Footnotereference then it will have footnote id value -->
					<xsl:if test="$noteClass='FootnoteReference'">
						<xsl:value-of select="concat('#footnote-',$noteID)"/>
					</xsl:if>
					<!--If Note_Class is Footnotereference then it will have footnote id value -->
					<xsl:if test="$noteClass='EndnoteReference'">
						<xsl:value-of select="concat('#endnote-',$noteID)"/>
					</xsl:if>
				</xsl:attribute>
				<!--Creating the attribute class for Noteref element and assinging it a value.-->
				<xsl:attribute name="class">
					<xsl:if test="$noteClass='FootnoteReference'">
						<xsl:value-of select="substring($noteClass,1,8)"/>
					</xsl:if>
					<!--Creating the attribute class for Noteref element and assinging it a value.-->
					<xsl:if test="$noteClass='EndnoteReference'">
						<xsl:value-of select="substring($noteClass,1,7)"/>
					</xsl:if>
				</xsl:attribute>
				<!--Checking for languages-->
				<xsl:if test="(w:rPr/w:lang) or (w:rPr/w:rFonts/@w:hint)">
					<xsl:attribute name="xml:lang">
						<xsl:call-template name="Languages">
							<xsl:with-param name="Attribute" select="true()"/>
						</xsl:call-template>
					</xsl:attribute>
				</xsl:if>
				<xsl:value-of select="$noteID"/>
			</noteref>
		</xsl:if>
	</xsl:template>

	<!--Template for Adding footnote-->
	<xsl:template name="InsertFootnotes">
		<xsl:param name="level"/>
		<xsl:param name="verfoot" as="xs:string"/>
		<xsl:param name="characterStyle" as="xs:boolean" select="false()"/>
		<xsl:param name="sOperators" as="xs:string"/>
		<xsl:param name="sMinuses" as="xs:string"/>
		<xsl:param name="sNumbers" as="xs:string"/>
		<xsl:param name="sZeros" as="xs:string"/>
		<!--Inserting default footnote id in the array list-->
		<xsl:variable name="checkid" as="xs:integer" select="d:FootNoteId($myObj,0, $level)"/>
		<!-- Checking for the matching Id and level returned from java code -->
		<xsl:if test="$checkid!=0">
			<!--Traversing through each footnote element in footnotes.xml file-->
			<xsl:for-each select="$footnotesXml//w:footnotes/w:footnote">
				<!--Checking if Id returned from C# is equal to the footnote Id in footnotes.xml file-->
				<xsl:if test="number(@w:id)=$checkid">
					<xsl:message terminate="no">progress:Insert footnote <xsl:value-of select="$checkid"/></xsl:message>
					<!--Creating note element and it's attribute values-->
					<note id="{concat('footnote-',$checkid)}" class="Footnote">
						<xsl:sequence select="d:sink(d:PushLevel($myObj, $level + 1))"/>
						<!--Travering each element inside w:footnote in footnote.xml file-->
						<xsl:for-each select="./node()">
							<!--Checking for Paragraph element-->
							<xsl:if test="self::w:p">
								<xsl:choose>
									<!--Checking for MathImage in Word2003/xp  footnotes-->
									<xsl:when test="(w:r/w:object/v:shape/v:imagedata/@r:id) and (not(w:r/w:object/o:OLEObject[@ProgID='Equation.DSMT4']))" >
										<p>
											<xsl:value-of select="$FootnotesNumberingPrefix"/>
											<xsl:choose>
												<xsl:when test="$FootnotesNumbering = 'number'">
													<xsl:value-of select="$checkid + number($FootnotesStartValue)"/>
												</xsl:when>
											</xsl:choose>
											<xsl:value-of select="$FootnotesNumberingSuffix"/>
											<imggroup>
												<img>
													<!--Variable to hold r:id from document.xml-->
													<xsl:variable name="Math_id"  as="xs:string" select="w:r/w:object/v:shape/v:imagedata/@r:id" />
													<xsl:choose>
														<!--Checking for alt text for MathEquation Image or providing 'Math Equation' as alttext-->
														<xsl:when test="w:r/w:object/v:shape/@alt">
															<xsl:sequence select="w:r/w:object/v:shape/@alt"/>
														</xsl:when>
														<xsl:otherwise>
															<xsl:attribute name="alt" select ="'Math Equation'"/>
														</xsl:otherwise>
													</xsl:choose>
													<!--Attribute holding the name of the Image-->
													<xsl:attribute name="src" select ="d:MathImageFootnote($myObj,$Math_id)">
														<!--Caling MathImageFootnote for copying Image to output folder-->
													</xsl:attribute>
												</img>
											</imggroup>
										</p>
									</xsl:when>
									<xsl:when test="w:r/w:object/o:OLEObject[@ProgID='Equation.DSMT4']">
										<xsl:variable name="Math_DSMT4" as="xs:string" select="d:GetMathML($myObj,'wdFootnotesStory')"/>
										<xsl:choose>
											<xsl:when test="$Math_DSMT4=''">
												<imggroup>
													<img>
														<!--Creating variable mathimage for storing r:id value from document.xml-->
														<xsl:variable name="Math_rid" as="xs:string" select="w:r/w:object/v:shape/v:imagedata/@r:id"/>
														<xsl:choose>
															<!--Checking for alt Text-->
															<xsl:when test="w:r/w:object/v:shape/@alt">
																<xsl:sequence select="w:r/w:object/v:shape/@alt"/>
															</xsl:when>
															<xsl:otherwise>
																<!--Hardcoding value 'Math Equation'if user donot provide alt text for Math Equations-->
																<xsl:attribute name="alt" select ="'Math Equation'"/>
															</xsl:otherwise>
														</xsl:choose>
														<xsl:attribute name="src" select ="d:MathImageFootnote($myObj,$Math_rid)">
															<!--Calling MathImage function-->
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
											<xsl:with-param name="flagNote" select="'footnote'"/>
											<xsl:with-param name="checkid" select="$checkid + 1"/>
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
						<xsl:sequence select="d:sink(d:PopLevel($myObj))"/>
					</note>
				</xsl:if>
				<xsl:sequence select="d:sink(d:InitializeNoteFlag($myObj))"/> <!-- empty -->
			</xsl:for-each>
			<!--Calling the template footnote recursively until the function returns 0-->
			<xsl:call-template name="InsertFootnotes">
				<xsl:with-param name="level" select="$level" />
				<xsl:with-param name="verfoot" select ="$verfoot"/>
				<xsl:with-param name="sOperators" select="$sOperators"/>
				<xsl:with-param name="sMinuses" select="$sMinuses"/>
				<xsl:with-param name="sNumbers" select="$sNumbers"/>
				<xsl:with-param name="sZeros" select="$sZeros"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<!--Template for handling multiple Prodnotes and Captions applied to an image-->
	<xsl:template name="ProcessCaptionProdNote">
		<xsl:param name="followingnodes" as="node()*"/>
		<xsl:param name="imageId" as="xs:string"/>
		<xsl:param name="characterStyle" as="xs:boolean"/>
		<xsl:choose>
			<!--Checking for inbuilt caption and Image-CaptionDAISY custom paragraph style-->
			<xsl:when test="($followingnodes[1]/w:pPr/w:pStyle/@w:val='Caption') or ($followingnodes[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
				<xsl:sequence select="d:sink(d:AddCaptionsProdnotes($myObj))"/> <!-- empty -->
				<caption>
					<!--attribute holds the value of the image id-->
					<xsl:attribute name="imgref" select="$imageId"/>
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
						<!--Variable holds the value which indicates that the image is bidirectionally oriented-->
						<xsl:variable name="Bd" as="xs:string">
							<!--calling the PictureLanguage template-->
							<xsl:call-template name="PictureLanguage">
								<xsl:with-param name="CheckLang" select="'picture'"/>
							</xsl:call-template>
						</xsl:variable>
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;bdo dir= &quot;rtl&quot; xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
					</xsl:if>
					<!--Looping through each of the node to print text to the output xml-->
					<xsl:for-each select="$followingnodes[1]/node()">
						<xsl:if test="self::w:r">
							<xsl:call-template name="TempCharacterStyle">
								<xsl:with-param name="characterStyle" select="$characterStyle"/>
							</xsl:call-template>
						</xsl:if>
						<xsl:if test="self::w:fldSimple">
							<xsl:value-of select="w:r/w:t"/>
						</xsl:if>
					</xsl:for-each>
					<!--Checking if image is bidirectionally oriented-->
					<xsl:if test="$followingnodes[1]/w:pPr/w:bidi">
						<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
						<xsl:value-of disable-output-escaping="yes" select="'&lt;/p&gt;'"/>
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
				<xsl:sequence select="d:sink(d:AddCaptionsProdnotes($myObj))"/> <!-- empty -->
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;prodnote render= &quot;optional&quot; imgref=&quot;',$imageId,'&quot;&gt;')"/>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="($followingnodes[1]/w:pPr/w:bidi) or ($followingnodes[1]/w:r/w:rPr/w:rtl)">
					<!--Variable holds the value which indicates that the image is bidirectionally oriented-->
					<xsl:variable name="Bd" as="xs:string">
						<!--calling the PictureLanguage template-->
						<xsl:call-template name="PictureLanguage">
							<xsl:with-param name="CheckLang" select="'picture'"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;bdo dir= &quot;rtl&quot; xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
				</xsl:if>
				<!--Looping through each of the node to print text to the output xml-->
				<xsl:for-each select="$followingnodes[1]/node()">
					<xsl:if test="self::w:r">
						<xsl:call-template name ="TempCharacterStyle">
							<xsl:with-param name ="characterStyle" select="$characterStyle"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:for-each>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="$followingnodes[1]/w:pPr/w:bidi">
					<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
					<xsl:value-of disable-output-escaping="yes" select="'&lt;/p&gt;'"/>
				</xsl:if>
				<xsl:value-of disable-output-escaping="yes" select="'&lt;/prodnote &gt;'"/>
				<!--Recursively calling the ProcessCaptionProdNote template till all the ProdNotes are processed-->
				<xsl:call-template name="ProcessCaptionProdNote">
					<xsl:with-param name="followingnodes" select="$followingnodes[position() > 1]"/>
					<xsl:with-param name="imageId" select="$imageId"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
			</xsl:when>
			<!--Checking for inbuilt caption and Prodnote-RequiredDAISY custom paragraph style-->
			<xsl:when test="($followingnodes[1]/w:pPr/w:pStyle/@w:val='Prodnote-RequiredDAISY')">
				<xsl:sequence select="d:sink(d:AddCaptionsProdnotes($myObj))"/> <!-- empty -->
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;prodnote render=&quot;required&quot; imgref=&quot;', $imageId ,'&quot;&gt;')"/>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="($followingnodes[1]/w:pPr/w:bidi) or ($followingnodes[1]/w:r/w:rPr/w:rtl)">
					<!--Variable holds the value which indicates that the image is bidirectionally oriented-->
					<xsl:variable name="Bd" as="xs:string">
						<!--calling the PictureLanguage template-->
						<xsl:call-template name="PictureLanguage">
							<xsl:with-param name="CheckLang" select="'picture'"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;bdo dir= &quot;rtl&quot; xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
				</xsl:if>
				<!--Looping through each of the node to print text to the output xml-->
				<xsl:for-each select="$followingnodes[1]/node()">
					<xsl:if test="self::w:r">
						<xsl:call-template name ="TempCharacterStyle">
							<xsl:with-param name ="characterStyle" select="$characterStyle"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:for-each>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="$followingnodes[1]/w:pPr/w:bidi">
					<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
					<xsl:value-of disable-output-escaping="yes" select="'&lt;/p&gt;'"/>
				</xsl:if>
				<xsl:value-of disable-output-escaping="yes" select="'&lt;/prodnote &gt;'"/>
				<!--Recursively calling the ProcessCaptionProdNote template till all the ProdNotes are processed-->
				<xsl:call-template name="ProcessCaptionProdNote">
					<xsl:with-param name="followingnodes" select="$followingnodes[position() > 1]"/>
					<xsl:with-param name="imageId" select="$imageId"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
			</xsl:when>
		</xsl:choose>
		<xsl:sequence select="d:ResetCaptionsProdnotes($myObj)"/> <!-- empty -->
	</xsl:template>

	<!--Template for implementing Simple Images i.e, ungrouped images-->
	<xsl:template name="PictureHandler">
		<xsl:param name="imgOpt" as="xs:string"/>
		<xsl:param name="dpi" as="xs:float?"/>
		<xsl:param name="characterStyle" as="xs:boolean"/>
		<xsl:message terminate="no">debug in PictureHandler</xsl:message>
		<xsl:variable name="alttext" as="xs:string?" select="w:drawing/wp:inline/wp:docPr/@descr"/>
		<!--Variable holds the value of Image Id-->
		<xsl:variable name="Img_Id" as="xs:string?">
			<xsl:choose>
				<xsl:when  test="w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed">
					<xsl:sequence select="w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed">
					<xsl:sequence select="w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:inline/wp:docPr/@id">
					<xsl:sequence select="w:drawing/wp:inline/wp:docPr/@id"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/wp:docPr/@id">
					<xsl:sequence select="w:drawing/wp:anchor/wp:docPr/@id"/>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<!--Variable holds the filename of the image-->
		<xsl:variable name="imageName" as="xs:string">
			<xsl:choose>
				<xsl:when  test="w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:nvPicPr/pic:cNvPr/@name">
					<xsl:sequence select="w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:nvPicPr/pic:cNvPr/@name"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic/pic:nvPicPr/pic:cNvPr/@name">
					<xsl:sequence select="w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic/pic:nvPicPr/pic:cNvPr/@name"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:sequence select="''"/>
				</xsl:otherwise>
			</xsl:choose>

		</xsl:variable>
		<!--Variable holds the value of Image Id concatenated with some random number generated for Image Id-->
		<xsl:variable name="imageId" as="xs:string">
			<xsl:choose>
				<xsl:when  test="w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed">
					<xsl:sequence select="concat($Img_Id,d:GenerateImageId($myObj))"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed">
					<xsl:sequence select="concat($Img_Id,d:GenerateImageId($myObj))"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:inline/wp:docPr/@id">
					<xsl:variable name="id" as="xs:string" select="../w:bookmarkStart[last()]/@w:name"/>
					<xsl:sequence select="d:CheckShapeId($myObj,$id)"/>
				</xsl:when>
				<xsl:when test="contains(w:drawing/wp:inline/wp:docPr/@name,'Diagram')">
					<xsl:sequence select="d:CheckShapeId($myObj,concat('Shape',substring-after(../../../../@id,'s')))"/>
				</xsl:when>
				<xsl:when test="contains(w:drawing/wp:inline/wp:docPr/@name,'Chart')">
					<xsl:sequence select="d:sink(d:CheckShapeId($myObj,concat('Shape',../w:bookmarkStart[last()]/@w:name)))"/> <!-- empty -->
					<xsl:sequence select="''"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:sequence select="''"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<!--
			imageWidth and imageHeight are expressed in EMU
			- 914400 EMU in an inch
			- 9525 EMU in a pixel @ 96 dpi
			- https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML
		-->
		<xsl:variable name="imageWidth" as="xs:double">
			<xsl:choose>
				<xsl:when  test="w:drawing/wp:inline/wp:extent">
					<xsl:sequence select="w:drawing/wp:inline/wp:extent/@cx"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/wp:extent">
					<xsl:sequence select="w:drawing/wp:anchor/wp:extent/@cx"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:inline/wp:extent">
					<xsl:sequence select="w:drawing/wp:inline/wp:extent/@cx"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/wp:extent">
					<xsl:sequence select="w:drawing/wp:anchor/wp:extent/@cx"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:sequence select="number(())"/> <!-- NaN -->
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="imageHeight" as="xs:double">
			<xsl:choose>
				<xsl:when  test="w:drawing/wp:inline/wp:extent">
					<xsl:sequence select="w:drawing/wp:inline/wp:extent/@cy"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/wp:extent">
					<xsl:sequence select="w:drawing/wp:anchor/wp:extent/@cy"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:inline/wp:extent">
					<xsl:sequence select="w:drawing/wp:inline/wp:extent/@cy"/>
				</xsl:when>
				<xsl:when test="w:drawing/wp:anchor/wp:extent">
					<xsl:sequence select="w:drawing/wp:anchor/wp:extent/@cy"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:sequence select="number(())"/> <!-- NaN -->
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<!--Checking if Img_Id variable contains any Image Id-->
		<xsl:if test="exists($Img_Id)">
			<!--Checking if document is bidirectionally oriented-->
			<xsl:if test="(../w:pPr/w:bidi) or (../w:pPr/w:jc/@w:val='right')">
				<!--Variable holds the value which indicates that the image is bidirectionally oriented-->
				<xsl:variable name="imgBd" as="xs:string">
					<!--calling the PictureLanguage template-->
					<xsl:call-template name="PictureLanguage">
						<xsl:with-param name="CheckLang" select="'picture'"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;bdo dir= &quot;rtl&quot; xml:lang=&quot;',$imgBd,'&quot;&gt;')"/>
			</xsl:if>
			<xsl:variable name="imageTest" as="xs:string">
				<xsl:choose>
					<xsl:when test="contains($Img_Id,'rId') and ($imgOpt='resize')">
						<xsl:sequence select ="d:Image($myObj,$Img_Id,$imageName)"/>
					</xsl:when>
					<xsl:when test="contains($Img_Id,'rId') and ($imgOpt='resample')">
						<xsl:sequence select ="d:ResampleImage($myObj,$Img_Id,$imageName,$dpi)"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<xsl:when test="contains($Img_Id,'rId')">
								<xsl:sequence select ="d:Image($myObj,$Img_Id,$imageName)"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:sequence select="concat($imageId,'.png')"/>
							</xsl:otherwise>
						</xsl:choose>

					</xsl:otherwise>
				</xsl:choose>
			</xsl:variable>
			<xsl:variable name="checkImage" as="xs:string" select="d:CheckImage($myObj,$imageTest)"/>
			<xsl:if test="$checkImage='1'">
				<!--Creating Imagegroup element-->
				<imggroup>
					<img>
						<!--attribute that holds the value of the Image ID-->
						<xsl:attribute name="id" select="$imageId"/>
						<!--attribute that holds the filename of the image returned for d:Image()-->
						<xsl:choose>
							<xsl:when test="$imgOpt='resize' and contains($Img_Id,'rId')">
								<xsl:attribute name="src" select ="$imageTest"/>
								<!--attribute that holds the alternate text for the image-->
								<xsl:attribute name="alt" select="$alttext"/>
								<xsl:attribute name="width" select="round(($imageWidth) div (9525))"/> <!-- assuming 96 dpi -->
								<xsl:attribute name="height" select="round(($imageHeight) div (9525))"/> <!-- assuming 96 dpi -->
							</xsl:when>
							<xsl:when test="$imgOpt='resample'  and contains($Img_Id,'rId')">
								<xsl:attribute name="src" select ="$imageTest"/>
								<!--attribute that holds the alternate text for the image-->
								<xsl:attribute name="alt" select="$alttext"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="src">
									<xsl:choose>
										<xsl:when test="contains($Img_Id,'rId')">
											<xsl:sequence select ="$imageTest"/>
										</xsl:when>
										<xsl:otherwise>
											<xsl:sequence select="$imageTest"/>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:attribute>
								<xsl:attribute name="alt" select="$alttext"/>
							</xsl:otherwise>
						</xsl:choose>
					</img>
					<!--Handling Image-CaptionDAISY custom paragraph style applied above an image-->
					<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY') or (../w:pPr/w:pStyle/@w:val='Caption') or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
						<caption>
							<xsl:attribute name="imgref" select="$imageId"/>
							<xsl:if test="(../following-sibling::w:p[1]/w:r/w:rPr/w:lang) or (../following-sibling::w:p[1]/w:r/w:rPr/w:rFonts/@w:hint)">
								<xsl:attribute name="xml:lang">
									<xsl:call-template name="PictureLanguage">
										<xsl:with-param name="CheckLang" select="'picture'"/>
									</xsl:call-template>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="(../following-sibling::w:p[1]/w:pPr/w:bidi) or (../following-sibling::w:p[1]/w:r/w:rPr/w:rtl)">
								<xsl:variable name="Bd" as="xs:string">
									<xsl:call-template name="PictureLanguage">
										<xsl:with-param name="CheckLang" select="'picture'"/>
									</xsl:call-template>
								</xsl:variable>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p  xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;bdo dir= &quot;rtl&quot; xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
							</xsl:if>
							<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
								<xsl:for-each select="../preceding-sibling::node()[1]/node()">
									<!--Printing the Caption value-->
									<xsl:if test="self::w:r">
										<xsl:call-template name ="TempCharacterStyle">
											<xsl:with-param name ="characterStyle" select="$characterStyle"/>
										</xsl:call-template>
									</xsl:if>
									<xsl:if test="self::w:fldSimple">
										<xsl:value-of select="w:r/w:t"/>
									</xsl:if>

								</xsl:for-each>
								<xsl:text> </xsl:text>
							</xsl:if>
							<xsl:if test="(../w:pPr/w:pStyle/@w:val='Caption') or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
								<xsl:for-each select="../node()">
									<!--Printing the Caption value-->
									<xsl:if test="self::w:r">
										<xsl:call-template name="TempCharacterStyle">
											<xsl:with-param name="characterStyle" select="$characterStyle"/>
										</xsl:call-template>
									</xsl:if>
									<xsl:if test="self::w:fldSimple">
										<xsl:value-of select="w:r/w:t"/>
									</xsl:if>

								</xsl:for-each>
								<xsl:text> </xsl:text>
							</xsl:if>
							<xsl:if test="../following-sibling::w:p[1]/w:pPr/w:bidi">
								<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
								<xsl:value-of disable-output-escaping="yes" select="'&lt;/p&gt;'"/>
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
					<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
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
		<xsl:param name="followingnodes" as="node()*"/>
		<xsl:param name="imageId" as="xs:string"/>
		<xsl:param name="characterStyle" as="xs:boolean"/>
		<xsl:choose>
			<!--Checking for Image-CaptionDAISY custom paragraph style-->
			<xsl:when test="($followingnodes[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
				<xsl:sequence select="d:sink(d:AddCaptionsProdnotes($myObj))"/> <!-- empty -->
				<caption>
					<xsl:attribute name="imgref" select="$imageId"/>
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
						<!--Variable holds the value which indicates that the image is bidirectionally oriented-->
						<xsl:variable name="Bd" as="xs:string">
							<!--calling the PictureLanguage template-->
							<xsl:call-template name="PictureLanguage">
								<xsl:with-param name="CheckLang" select="'imagegroup'"/>
							</xsl:call-template>
						</xsl:variable>
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
						<xsl:value-of disable-output-escaping="yes" select="concat('&lt;bdo dir= &quot;rtl&quot; xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
					</xsl:if>
					<!--Looping through each of the node to print the text to the output xml-->
					<xsl:for-each select="$followingnodes[1]/node()">
						<xsl:if test="self::w:r">
							<xsl:call-template name="TempCharacterStyle">
								<xsl:with-param name="characterStyle" select="$characterStyle"/>
							</xsl:call-template>
						</xsl:if>
						<xsl:if test="self::w:fldSimple">
							<xsl:value-of select="w:r/w:t"/>
						</xsl:if>

					</xsl:for-each>
					<!--Checking for image is bidirectionally oriented-->
					<xsl:if test="$followingnodes[1]/w:pPr/w:bidi">
						<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
						<xsl:value-of disable-output-escaping="yes" select="'&lt;/p&gt;'"/>
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
				<xsl:sequence select="d:sink(d:AddCaptionsProdnotes($myObj))"/> <!-- empty -->
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;prodnote render=&quot;optional&quot; imgref=&quot;',$imageId,'&quot;&gt;')"/>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="($followingnodes[1]/w:pPr/w:bidi) or ($followingnodes[1]/w:r/w:rPr/w:rtl)">
					<!--Variable holds the value which indicates that the image is bidirectionally oriented-->
					<xsl:variable name="Bd" as="xs:string">
						<!--Calling the PictureLanguage template-->
						<xsl:call-template name="PictureLanguage">
							<xsl:with-param name="CheckLang" select="'imagegroup'"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;bdo dir= &quot;rtl&quot; xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
				</xsl:if>
				<!--Looping through each of the node to print the text to the output xml-->
				<xsl:for-each select="$followingnodes[1]/node()">
					<xsl:if test="self::w:r">
						<xsl:call-template name ="TempCharacterStyle">
							<xsl:with-param name ="characterStyle" select="$characterStyle"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:for-each>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="$followingnodes[1]/w:pPr/w:bidi">
					<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
					<xsl:value-of disable-output-escaping="yes" select="'&lt;/p&gt;'"/>
				</xsl:if>
				<xsl:value-of disable-output-escaping="yes" select="'&lt;/prodnote &gt;'"/>
				<!--Recursively calling the ProcessCaptionProdNote template till all the prodnotes are processed-->
				<xsl:call-template name="ProcessProdNoteImggroups">
					<xsl:with-param name="followingnodes" select="$followingnodes[position() > 1]"/>
					<xsl:with-param name="imageId" select="$imageId"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
			</xsl:when>
			<!--Checking for Prodnote-RequiredDAISY custom paragraph style-->
			<xsl:when test="($followingnodes[1]/w:pPr/w:pStyle/@w:val='Prodnote-RequiredDAISY')">
				<xsl:sequence select="d:sink(d:AddCaptionsProdnotes($myObj))"/> <!-- empty -->
				<xsl:value-of disable-output-escaping="yes" select="concat('&lt;prodnote render=&quot;required&quot; imgref=&quot;',$imageId,'&quot;&gt;')"/>
				<!--Getting the language id by calling the PictureLanguage template-->
				<xsl:if test="($followingnodes[1]/w:pPr/w:bidi) or ($followingnodes[1]/w:r/w:rPr/w:rtl)">
					<!--attribute that holds language id-->
					<xsl:variable name="Bd" as="xs:string">
						<!--calling the PictureLanguage template-->
						<xsl:call-template name="PictureLanguage">
							<xsl:with-param name="CheckLang" select="'imagegroup'"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
					<xsl:value-of disable-output-escaping="yes" select="concat('&lt;bdo dir= &quot;rtl&quot; xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
				</xsl:if>
				<!--Looping through each of the node to print the text to the output xml-->
				<xsl:for-each select="$followingnodes[1]/node()">
					<xsl:if test="self::w:r">
						<xsl:call-template name="TempCharacterStyle">
							<xsl:with-param name="characterStyle" select="$characterStyle"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:for-each>
				<!--Checking if image is bidirectionally oriented-->
				<xsl:if test="$followingnodes[1]/w:pPr/w:bidi">
					<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
					<xsl:value-of disable-output-escaping="yes" select="'&lt;/p&gt;'"/>
				</xsl:if>
				<xsl:value-of disable-output-escaping="yes" select="'&lt;/prodnote &gt;'"/>
				<!--Recursively calling the ProcessCaptionProdNote template till all the prodnotes are processed-->
				<xsl:call-template name="ProcessProdNoteImggroups">
					<xsl:with-param name="followingnodes" select="$followingnodes[position() > 1]"/>
					<xsl:with-param name="imageId" select="$imageId"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
			</xsl:when>
		</xsl:choose>
		<xsl:sequence select="d:ResetCaptionsProdnotes($myObj)"/> <!-- empty -->
	</xsl:template>

	<!--Template for Implementing grouped images-->
	<xsl:template name="Imagegroups">
		<xsl:param name="characterStyle" as="xs:boolean"/>
		<!--Handling Image-CaptionDAISY custom paragraph style applied above an image-->
		<xsl:if test="../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY'">
			<xsl:variable name="caption" as="xs:string*">
				<xsl:for-each select="../preceding-sibling::node()[1]/node()">
					<xsl:if test="self::w:r">
						<xsl:call-template name ="TempCharacterStyle">
							<xsl:with-param name ="characterStyle" select="$characterStyle"/>
						</xsl:call-template>
					</xsl:if>
					<xsl:if test="self::w:fldSimple">
						<xsl:sequence select="w:r/w:t"/>
					</xsl:if>

				</xsl:for-each>
			</xsl:variable>
			<xsl:variable name="caption" as="xs:string" select="string-join($caption,'')"/>
			<xsl:sequence select="d:sink(d:InsertCaption($myObj,$caption))"/> <!-- empty -->
		</xsl:if>
		<!--Looping through each pict element and storing the caption value in the caption variable-->

		<xsl:if test="../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:pPr/w:pStyle[@w:val='Caption']">
			<xsl:variable name="caption" as="xs:string*">
				<xsl:for-each select="../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/node()">
					<xsl:if test="self::w:r">
						<xsl:call-template name="TempCharacterStyle">
							<xsl:with-param name="characterStyle" select="$characterStyle"/>
						</xsl:call-template>
					</xsl:if>
					<xsl:if test="self::w:fldSimple">
						<xsl:sequence select="w:r/w:t"/>
					</xsl:if>

				</xsl:for-each>
			</xsl:variable>
			<xsl:variable name="caption" as="xs:string" select="string-join($caption,'')"/>
			<!--Inserting the caption value in the Arraylist-->
			<xsl:sequence select="d:sink(d:InsertCaption($myObj,$caption))"/> <!-- empty -->
		</xsl:if>
		<xsl:variable name="Imageid" as="xs:string" select ="d:CheckShapeId($myObj,concat('Shape',substring-after(w:pict/v:group/@id,'s')))"/>
		<xsl:variable name="checkImage" as="xs:string" select="d:CheckImage($myObj,concat($Imageid,'.png'))"/>
		<xsl:if test="$checkImage='1'">
			<!--Checking for the presence of Images-->
			<imggroup>
				<img>
					<!--Creating attribute id of img element-->
					<xsl:attribute name="id" select="$Imageid"/>
					<!--Creating attribute alt for alternate text of img element-->
					<xsl:attribute name="alt" select="w:pict/v:group/@alt"/>
					<!--Creating attribute src of img element-->
					<xsl:attribute name="src" select ="concat($Imageid,'.png')"/>
				</img>

				<xsl:variable name="checkcaption" as="xs:string" select="d:ReturnCaption($myObj)"/>
				<!--Checking if checkcaption variables holds any value-->
				<xsl:if test="$checkcaption!='0'">
					<caption>
						<xsl:attribute name="imgref" select="$Imageid"/>
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
							<xsl:variable name="Bd" as="xs:string">
								<xsl:call-template name="PictureLanguage">
									<xsl:with-param name="CheckLang" select="'imagegroup'"/>
								</xsl:call-template>
							</xsl:variable>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;bdo dir= &quot;rtl&quot; xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
						</xsl:if>
						<xsl:value-of select="$checkcaption"/>
						<xsl:if test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:pPr/w:bidi">
							<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
							<xsl:value-of disable-output-escaping="yes" select="'&lt;/p&gt;'"/>
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
		<xsl:param name="characterStyle" as="xs:boolean"/>
		<!--Variable that holds the Image Id-->
		<xsl:variable name="imageId" as="xs:string" select="concat(w:pict/v:shape/v:imagedata/@r:id,d:GenerateImageId($myObj))"/>
		<!--Checking if image is bidirectionally oriented-->
		<xsl:if test="(../w:pPr/w:bidi) or (../w:pPr/w:jc/@w:val='right')">
			<xsl:variable name="imgBd" as="xs:string">
				<xsl:call-template name="PictureLanguage">
					<xsl:with-param name="CheckLang" select="'picture'"/>
				</xsl:call-template>
			</xsl:variable>
			<xsl:value-of disable-output-escaping="yes" select="concat('&lt;bdo dir= &quot;rtl&quot; xml:lang=&quot;',$imgBd,'&quot;&gt;')"/>
		</xsl:if>
		<xsl:variable name="checkImage" as="xs:string" select="d:CheckImage($myObj,d:Image($myObj,w:pict/v:shape/v:imagedata/@r:id,w:pict/v:shape/v:imagedata/@o:title))"/>
		<xsl:if test="$checkImage='1'">
			<imggroup>
				<img>
					<!--attribute to store Image id-->
					<xsl:attribute name="id" select="$imageId"/>
					<!--variable to store Image name-->
					<xsl:variable name="image2003Name" as="xs:string" select="w:pict/v:shape/v:imagedata/@o:title"/>
					<!--variable to store Image id-->
					<xsl:variable name="rid" as="xs:string" select="w:pict/v:shape/v:imagedata/@r:id"/>
					<!--Creating attribute src of img element-->
					<xsl:attribute name="src" select ="d:Image($myObj,$rid,$image2003Name)"/>
					<!--Creating attribute alt for alternate text of img element-->
					<xsl:attribute name="alt" select="w:pict/v:shape/@alt"/>
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
							<xsl:variable name="Bd" as="xs:string">
								<xsl:call-template name="PictureLanguage">
									<xsl:with-param name="CheckLang" select="'picture'"/>
								</xsl:call-template>
							</xsl:variable>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p  xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;bdo dir= &quot;rtl&quot; xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
						</xsl:if>
						<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
							<xsl:for-each select="../preceding-sibling::node()[1]/node()">
								<!--Printing the Caption value-->
								<xsl:if test="self::w:r">
									<xsl:call-template name="TempCharacterStyle">
										<xsl:with-param name="characterStyle" select="$characterStyle"/>
									</xsl:call-template>
								</xsl:if>
								<xsl:if test="self::w:fldSimple">
									<xsl:value-of select="w:r/w:t"/>
								</xsl:if>

							</xsl:for-each>
							<xsl:text> </xsl:text>
						</xsl:if>
						<xsl:if test="(../w:pPr/w:pStyle/@w:val='Caption') or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
							<xsl:for-each select="../node()">
								<!--Printing the Caption value-->
								<xsl:if test="self::w:r">
									<xsl:call-template name="TempCharacterStyle">
										<xsl:with-param name="characterStyle" select="$characterStyle"/>
									</xsl:call-template>
								</xsl:if>
								<xsl:if test="self::w:fldSimple">
									<xsl:value-of select="w:r/w:t"/>
								</xsl:if>

							</xsl:for-each>
							<xsl:text> </xsl:text>
						</xsl:if>
						<xsl:if test="../following-sibling::w:p[1]/w:pPr/w:bidi">
							<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
							<xsl:value-of disable-output-escaping="yes" select="'&lt;/p&gt;'"/>
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
				<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
			</xsl:if>
		</xsl:if>
		<xsl:if test="$checkImage='0'">
			<xsl:message terminate="no">translation.oox2Daisy.Image</xsl:message>
		</xsl:if>
	</xsl:template>


	<xsl:template name="Object">
		<xsl:param name="characterStyle" as="xs:boolean"/>
		<xsl:if test="not(contains(w:object/o:OLEObject/@ProgID,'Equation'))">
			<xsl:if test="(
				contains(w:object/o:OLEObject/@ProgID,'Excel')
				or contains(w:object/o:OLEObject/@ProgID,'Word')
				or contains(w:object/o:OLEObject/@ProgID,'PowerPoint')
			)">
				<xsl:variable name="href" as="xs:string" select="d:Object($myObj,w:object/o:OLEObject/@r:id)"/>
				<xsl:text disable-output-escaping="yes">&lt;a href=&quot;</xsl:text>
				<xsl:value-of select="$href"/>
				<xsl:text disable-output-escaping="yes">&quot; external=&quot;true&quot;&gt;</xsl:text>
				<!--<xsl:value-of disable-output-escaping="yes" select="concat('&lt;a href=&quot;',$href,'&quot; external=&quot;true&quot;&gt;')"/>-->
			</xsl:if>
			<xsl:variable name="ImageName" as="xs:string" select="d:MathImage($myObj,w:object/v:shape/v:imagedata/@r:id)"/>
			<xsl:variable name ="id" as="xs:string" select="d:GenerateObjectId($myObj)"/>
			<xsl:variable name="ImageId" as="xs:string" select="concat($ImageName,$id)"/>
			<xsl:variable name="checkImage" as="xs:string" select="d:CheckImage($myObj,$ImageName)"/>
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
								<xsl:variable name="Bd" as="xs:string">
									<xsl:call-template name="PictureLanguage">
										<xsl:with-param name="CheckLang" select="'picture'"/>
									</xsl:call-template>
								</xsl:variable>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;bdo dir= &quot;rtl&quot; xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
							</xsl:if>
							<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
								<xsl:for-each select="../preceding-sibling::node()[1]/node()">
									<!--Printing the Caption value-->
									<xsl:if test="self::w:r">
										<xsl:call-template name="TempCharacterStyle">
											<xsl:with-param name="characterStyle" select="$characterStyle"/>
										</xsl:call-template>
									</xsl:if>
									<xsl:if test="self::w:fldSimple">
										<xsl:value-of select="w:r/w:t"/>
									</xsl:if>

								</xsl:for-each>
							</xsl:if>
							<xsl:if test="(../w:pPr/w:pStyle/@w:val='Caption') or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
								<xsl:for-each select="../node()">
									<!--Printing the Caption value-->
									<xsl:if test="self::w:r">
										<xsl:call-template name="TempCharacterStyle">
											<xsl:with-param name="characterStyle" select="$characterStyle"/>
										</xsl:call-template>
									</xsl:if>
									<xsl:if test="self::w:fldSimple">
										<xsl:value-of select="w:r/w:t"/>
									</xsl:if>

								</xsl:for-each>
								<xsl:text> </xsl:text>
							</xsl:if>
							<xsl:if test="../following-sibling::w:p[1]/w:pPr/w:bidi">
								<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
								<xsl:value-of disable-output-escaping="yes" select="'&lt;/p&gt;'"/>
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
					<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
				</xsl:if>
				<xsl:if test="contains(w:object/o:OLEObject/@ProgID,'Excel') or contains(w:object/o:OLEObject/@ProgID,'Word') or contains(w:object/o:OLEObject/@ProgID,'PowerPoint')">
					<xsl:value-of disable-output-escaping="yes" select="'&lt;/a&gt;'"/>
				</xsl:if>
			</xsl:if>
			<xsl:if test="$checkImage='0'">
				<xsl:message terminate="no">translation.oox2Daisy.Image</xsl:message>
			</xsl:if>
		</xsl:if>
	</xsl:template>

	<xsl:template name="tmpShape">
		<xsl:param name="characterStyle" as="xs:boolean"/>
		<xsl:variable name="imageId" as="xs:string">
			<xsl:choose>
				<xsl:when test="(w:pict/v:shape/@id) and (w:pict/v:shape/@o:spid)">
					<xsl:sequence select="d:CheckShapeId($myObj,concat('Shape',substring-after(w:pict/v:shape/@o:spid,'s')))"/>
				</xsl:when>
				<xsl:when test="w:pict/v:shape/@id">
					<xsl:sequence select="d:CheckShapeId($myObj,concat('Shape',substring-after(w:pict/v:shape/@id,'s')))"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:sequence select="d:CheckShapeId($myObj,concat('Shape',substring-after(w:pict//@id,'s')))"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="checkImage" as="xs:string" select="d:CheckImage($myObj,concat($imageId,'.png'))"/>
		<xsl:if test="$checkImage='1'">
			<imggroup>
				<img>
					<xsl:attribute name="id" select="$imageId"/>
					<xsl:attribute name="src" select="concat($imageId,'.png')"/>
					<xsl:attribute name="alt" select="w:pict/v:shape/@alt"/>
				</img>

				<xsl:if test="(
					(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY') 
					or (../w:pPr/w:pStyle/@w:val='Caption') 
					or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')
				)">
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
							<xsl:variable name="Bd" as="xs:string">
								<xsl:call-template name="PictureLanguage">
									<xsl:with-param name="CheckLang" select="'picture'"/>
								</xsl:call-template>
							</xsl:variable>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p  xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
							<xsl:value-of disable-output-escaping="yes" select="concat('&lt;bdo dir= &quot;rtl&quot; xml:lang=&quot;',$Bd,'&quot;&gt;')"/>
						</xsl:if>
						<xsl:if test="(../preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
							<xsl:for-each select="../preceding-sibling::node()[1]/node()">

								<!--Printing the Caption value-->

								<xsl:if test="self::w:r">
									<xsl:call-template name="TempCharacterStyle">
										<xsl:with-param name="characterStyle" select="$characterStyle"/>
									</xsl:call-template>
								</xsl:if>
								<xsl:if test="self::w:fldSimple">
									<xsl:value-of select="w:r/w:t"/>
								</xsl:if>

							</xsl:for-each>
							<xsl:text> </xsl:text>
						</xsl:if>
						<xsl:if test="(../w:pPr/w:pStyle/@w:val='Caption') or (../w:pPr/w:pStyle/@w:val='Image-CaptionDAISY')">
							<xsl:for-each select="../node()">

								<!--Printing the Caption value-->


								<xsl:if test="self::w:r">
									<xsl:call-template name="TempCharacterStyle">
										<xsl:with-param name="characterStyle" select="$characterStyle"/>
									</xsl:call-template>
								</xsl:if>
								<xsl:if test="self::w:fldSimple">
									<xsl:value-of select="w:r/w:t"/>
								</xsl:if>

							</xsl:for-each>
							<xsl:text> </xsl:text>
						</xsl:if>
						<xsl:if test="../following-sibling::w:p[1]/w:pPr/w:bidi">
							<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
							<xsl:value-of disable-output-escaping="yes" select="'&lt;/p&gt;'"/>
						</xsl:if>
					</caption>
					<xsl:if test="../following-sibling::w:p[1]/w:pPr/w:bidi">
						<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
						<xsl:value-of disable-output-escaping="yes" select="'&lt;/p&gt;'"/>
					</xsl:if>
				</xsl:if>
				<xsl:call-template name="ProcessCaptionProdNote">
					<xsl:with-param name="followingnodes" select="../following-sibling::node()"/>
					<xsl:with-param name="imageId" select="$imageId"/>
					<xsl:with-param name="characterStyle" select="$characterStyle"/>
				</xsl:call-template>
			</imggroup>
			<xsl:if test="(../w:pPr/w:bidi) or (../w:pPr/w:jc/@w:val='right')">
				<xsl:value-of disable-output-escaping="yes" select="'&lt;/bdo&gt;'"/>
			</xsl:if>
		</xsl:if>
		<xsl:if test="$checkImage='0'">
			<xsl:message terminate="no">translation.oox2Daisy.Image</xsl:message>
		</xsl:if>
	</xsl:template>

	<!--Template for checking section breaks for page numbers-->
	<xsl:template name="SectionBreak">
		<xsl:param name="count" as="xs:integer"/>
		<xsl:param name="node" as="xs:string"/>
		<xsl:sequence select="d:sink(d:InitalizeCheckSectionBody($myObj))"/> <!-- empty -->
		<xsl:sequence select="d:sink(d:ResetSetConPageBreak($myObj))"/> <!-- empty -->
		<xsl:choose>
			<!--if page number for front matter-->
			<xsl:when test="$node='front'">
				<!--incrementing the default page counter-->
				<xsl:sequence select="d:sink(d:IncrementPage($myObj))"/> <!-- empty -->
				<!--Traversing through each node-->
				<xsl:for-each select="following-sibling::node()">
					<xsl:choose>
						<!--Checking for paragraph section break-->
						<xsl:when test="w:pPr/w:sectPr">

							<xsl:if test="d:CheckSectionBody($myObj)=1">
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
						<xsl:when test="self::w:sectPr">
							<xsl:if test="d:CheckSectionBody($myObj)=1">
								<xsl:sequence select="d:sink(d:CheckSectionFront($myObj))"/> <!-- empty -->
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
					<xsl:sequence select="d:sink(d:SetConPageBreak($myObj))"/> <!-- empty -->
				</xsl:if>
				<!--Traversing through each node-->
				<xsl:for-each select="../following-sibling::node()">
					<xsl:choose>
						<!--Checking for paragraph section break-->
						<xsl:when test="w:pPr/w:sectPr">
							<xsl:if test="d:CheckSectionBody($myObj)=1">
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
											<xsl:with-param name="counter" select="0"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page start is present and not page format-->
									<xsl:when test="not(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--If both are not present-->
									<xsl:otherwise>
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="0"/>
										</xsl:call-template>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:if>
						</xsl:when>
						<!--Checking for Section in a document-->
						<xsl:when test="self::w:sectPr">
							<xsl:if test="d:CheckSectionBody($myObj)=1">
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
											<xsl:with-param name="counter" select="0"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page start is present and not page format-->
									<xsl:when test="not(w:pgNumType/@w:fmt) and (w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--If both are not present-->
									<xsl:otherwise>
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="0"/>
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
					<xsl:sequence select="d:sink(d:SetConPageBreak($myObj))"/> <!-- empty -->
				</xsl:if>
				<!--Traversing through each node-->
				<xsl:for-each select="following-sibling::node()">
					<xsl:choose>
						<!--Checking for paragraph section break-->
						<xsl:when test="w:pPr/w:sectPr">
							<xsl:if test="d:CheckSectionBody($myObj)=1">
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
											<xsl:with-param name="counter" select="0"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page start is present and not page format-->
									<xsl:when test="not(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--If both are not present-->
									<xsl:otherwise>
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="0"/>
										</xsl:call-template>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:if>
						</xsl:when>
						<!--Checking for Section in a document-->
						<xsl:when test="self::w:sectPr">
							<xsl:if test="d:CheckSectionBody($myObj)=1">
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
											<xsl:with-param name="counter" select="0"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page start is present and not page format-->
									<xsl:when test="not(w:pgNumType/@w:fmt) and (w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--If both are not present-->
									<xsl:otherwise>
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
											<xsl:with-param name="matter" select="$node"/>
											<xsl:with-param name="counter" select="0"/>
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
					<xsl:sequence select="d:sink(d:SetConPageBreak($myObj))"/> <!-- empty -->
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
									<xsl:with-param name="counter" select="0"/>
								</xsl:call-template>
							</xsl:when>
							<!--Checking if page start is present and not page format-->
							<xsl:when test="not(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
								<!--Calling template for page number text-->
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
									<xsl:with-param name="matter" select="$node"/>
									<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
								</xsl:call-template>
							</xsl:when>
							<!--If both are not present-->
							<xsl:otherwise>
								<!--Calling template for page number text-->
								<xsl:call-template name="PageNumber">
									<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
									<xsl:with-param name="matter" select="$node"/>
									<xsl:with-param name="counter" select="0"/>
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
					<xsl:sequence select="d:sink(d:SetConPageBreak($myObj))"/> <!-- empty -->
				</xsl:if>
				<!--Traversing through each node-->
				<xsl:for-each select="../../following-sibling::node()">
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
										<xsl:with-param name="counter" select="0"/>
									</xsl:call-template>
								</xsl:when>
								<!--Checking if page start is present and not page format-->
								<xsl:when test="not(w:pPr/w:sectPr/w:pgNumType/@w:fmt) and (w:pPr/w:sectPr/w:pgNumType/@w:start)">
									<!--Calling template for page number text-->
									<xsl:call-template name="PageNumber">
										<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
										<xsl:with-param name="matter" select="$node"/>
										<xsl:with-param name="counter" select="w:pPr/w:sectPr/w:pgNumType/@w:start"/>
									</xsl:call-template>
								</xsl:when>
								<!--If both are not present-->
								<xsl:otherwise>
									<!--Calling template for page number text-->
									<xsl:call-template name="PageNumber">
										<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
										<xsl:with-param name="matter" select="$node"/>
										<xsl:with-param name="counter" select="0"/>
									</xsl:call-template>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<!--Checking for Section in a document-->
						<xsl:when test="self::w:sectPr">
							<xsl:if test="d:CheckSectionBody($myObj)=1">
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
											<xsl:with-param name="counter" select="0"/>
										</xsl:call-template>
									</xsl:when>
									<!--Checking if page start is present and not page format-->
									<xsl:when test="not(w:pgNumType/@w:fmt) and (w:pgNumType/@w:start)">
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
											<xsl:with-param name="matter" select="'body'"/>
											<xsl:with-param name="counter" select="w:pgNumType/@w:start"/>
										</xsl:call-template>
									</xsl:when>
									<!--If both are not present-->
									<xsl:otherwise>
										<!--Calling template for page number text-->
										<xsl:call-template name="PageNumber">
											<xsl:with-param name="pagetype" select="d:GetPageFormat($myObj)"/>
											<xsl:with-param name="matter" select="'body'"/>
											<xsl:with-param name="counter" select="0"/>
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
		<xsl:message terminate="no">debug in countpageTOC</xsl:message>
		<xsl:for-each select="preceding-sibling::*">
			<xsl:choose>
				<!--Checking for page break in TOC-->
				<xsl:when test="(w:r/w:br/@w:type='page') or (w:r/w:lastRenderedPageBreak)">
					<xsl:sequence select="d:sink(d:PageForTOC($myObj))"/> <!-- empty -->
					<xsl:sequence select="d:sink(d:IncrementPage($myObj))"/> <!-- empty -->
					<xsl:if test="not(w:r/w:t)">
						<!--Calling template for initializing page number info-->
						<xsl:call-template name="SectionBreak">
							<xsl:with-param name="count" select="d:ReturnPageNum($myObj)"/>
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
					<xsl:sequence select="d:sink(d:PageForTOC($myObj))"/> <!-- empty -->
					<xsl:sequence select="d:sink(d:IncrementPage($myObj))"/> <!-- empty -->
				</xsl:when>
			</xsl:choose>
		</xsl:for-each>
		<xsl:variable name="countPage" as="xs:integer" select="d:PageForTOC($myObj) - 1"/>
		<xsl:call-template name="SectionBreak">
			<xsl:with-param name="count" select="$countPage"/>
			<xsl:with-param name="node" select="'front'"/>
		</xsl:call-template>
	</xsl:template>

	<!--Template to translate page number information-->
	<xsl:template name="PageNumber">
		<xsl:param name="pagetype" as="xs:string"/>
		<xsl:param name="matter" as="xs:string"/>
		<xsl:param name="counter" as="xs:integer"/>
		<xsl:message terminate="no">debug in PageNumber</xsl:message>
		<xsl:choose>
			<xsl:when test="d:GetCurrentMatterType($myObj)='Frontmatter'">
				<xsl:if test="not((d:SetConPageBreak($myObj)&gt;1) and (w:type/@w:val='continuous'))">
					<xsl:variable name="count" as="xs:integer" select="d:IncrementPageNo($myObj)-1"/>
					<xsl:choose>
						<!--LowerRoman page number-->
						<xsl:when test="$pagetype='lowerRoman'">
							<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumLowerRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperRoman page number-->
						<xsl:when test="$pagetype='upperRoman'">
							<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumUpperRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--LowerLetter page number-->
						<xsl:when test="$pagetype='lowerLetter'">
							<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumLowerAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperLetter page number-->
						<xsl:when test="$pagetype='upperLetter'">
							<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumUpperAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--Page number with dash-->
						<xsl:when test="$pagetype='numberInDash'">
							<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="concat('-',$count,'-')"/>
							</pagenum>
						</xsl:when>
						<!--Normal page number-->
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test="$counter=0 and d:GetSectionFront($myObj)=1">
									<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
										<xsl:value-of select="$count"/>
									</pagenum>
								</xsl:when>
								<xsl:when test="$counter=0 and d:GetSectionFront($myObj)=0">
									<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
										<xsl:value-of select="d:ReturnPageNum($myObj)"/>
									</pagenum>
								</xsl:when>
								<xsl:otherwise>
									<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
										<xsl:value-of select="$count"/>
									</pagenum>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>
			</xsl:when>
	  
			<xsl:when test="d:GetCurrentMatterType($myObj)='Bodymatter'">
				<xsl:if test="not((d:SetConPageBreak($myObj)&gt;1) and (w:type/@w:val='continuous'))">
					<xsl:variable name="count" as="xs:integer" select="d:IncrementPageNo($myObj)-1"/>
		  			<xsl:choose>
						<!--LowerRoman page number-->
						<xsl:when test="$pagetype='lowerRoman'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumLowerRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperRoman page number-->
						<xsl:when test="$pagetype='upperRoman'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumUpperRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--LowerLetter page number-->
						<xsl:when test="$pagetype='lowerLetter'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumLowerAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperLetter page number-->
						<xsl:when test="$pagetype='upperLetter'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumUpperAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--Page number with dash-->
						<xsl:when test="$pagetype='numberInDash'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="concat('-',$count,'-')"/>
							</pagenum>
						</xsl:when>
						<!--Normal page number-->
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test="$counter=0 and d:GetSectionFront($myObj)=1">
									<pagenum page="normal" id="{concat('page',d:GeneratePageId($myObj))}">
										<xsl:value-of select="$count"/>
									</pagenum>
								</xsl:when>
								<xsl:when test="$counter=0 and d:GetSectionFront($myObj)=0">
									<pagenum page="normal" id="{concat('page',d:GeneratePageId($myObj))}">
										<xsl:value-of select="d:ReturnPageNum($myObj)"/>
									</pagenum>
								</xsl:when>
								<xsl:otherwise>
									<pagenum page="normal" id="{concat('page',d:GeneratePageId($myObj))}">
										<xsl:value-of select="$count"/>
									</pagenum>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>
			</xsl:when>

			<xsl:when test="d:GetCurrentMatterType($myObj)='Reartmatter'">
				<xsl:if test="not((d:SetConPageBreak($myObj)&gt;1) and (w:type/@w:val='continuous'))">
					<xsl:variable name="count" as="xs:integer" select="d:IncrementPageNo($myObj)-1"/>
					<xsl:choose>
						<!--LowerRoman page number-->
						<xsl:when test="$pagetype='lowerRoman'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumLowerRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperRoman page number-->
						<xsl:when test="$pagetype='upperRoman'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumUpperRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--LowerLetter page number-->
						<xsl:when test="$pagetype='lowerLetter'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumLowerAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperLetter page number-->
						<xsl:when test="$pagetype='upperLetter'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumUpperAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--Page number with dash-->
						<xsl:when test="$pagetype='numberInDash'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="concat('-',$count,'-')"/>
							</pagenum>
						</xsl:when>
						<!--Normal page number-->
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test="$counter=0 and d:GetSectionFront($myObj)=1">
									<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
										<xsl:value-of select="$count"/>
									</pagenum>
								</xsl:when>
								<xsl:when test="$counter=0 and d:GetSectionFront($myObj)=0">
									<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
										<xsl:value-of select="d:ReturnPageNum($myObj)"/>
									</pagenum>
								</xsl:when>
								<xsl:otherwise>
									<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
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
				<xsl:if test="d:GetSectionFront($myObj)=1">
					<xsl:sequence select="d:sink(d:IncrementPageNo($myObj))"/> <!-- empty -->
				</xsl:if>
				<xsl:choose>
					<!--LowerRoman page number-->
					<xsl:when test="$pagetype='lowerRoman'">
						<xsl:variable name="pageno" as="xs:string" select="d:PageNumLowerRoman($counter)"/>
						<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
							<xsl:value-of select="$pageno"/>
						</pagenum>
					</xsl:when>
					<!--UpperRoman page number-->
					<xsl:when test="$pagetype='upperRoman'">
						<xsl:variable name="pageno" as="xs:string" select="d:PageNumUpperRoman($counter)"/>
						<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
							<xsl:value-of select="$pageno"/>
						</pagenum>
					</xsl:when>
					<!--LowerLetter page number-->
					<xsl:when test="$pagetype='lowerLetter'">
						<xsl:variable name="pageno" as="xs:string" select="d:PageNumLowerAlphabet($counter)"/>
						<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
							<xsl:value-of select="$pageno"/>
						</pagenum>
					</xsl:when>
					<!--UpperLetter page number-->
					<xsl:when test="$pagetype='upperLetter'">
						<xsl:variable name="pageno" as="xs:string" select="d:PageNumUpperAlphabet($counter)"/>
						<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
							<xsl:value-of select="$pageno"/>
						</pagenum>
					</xsl:when>
					<!--Page number with dash-->
					<xsl:when test="$pagetype='numberInDash'">
						<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
							<xsl:value-of select="concat('-',$counter,'-')"/>
						</pagenum>
					</xsl:when>
					<!--Normal page number-->
					<xsl:otherwise>
						<pagenum page="front" id="{concat('page',d:GeneratePageId($myObj))}">
							<xsl:value-of select="$counter"/>
						</pagenum>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<!--Bodymatter page number-->
			<xsl:when test="($matter='body') or ($matter='bodysection') or ($matter='Para')">
				<xsl:if test="not((d:SetConPageBreak($myObj)&gt;1) and (w:type/@w:val='continuous'))">
					<xsl:variable name="count" as="xs:integer" select="d:IncrementPageNo($myObj)-1"/>
					<xsl:choose>
						<!--LowerRoman page number-->
						<xsl:when test="$pagetype='lowerRoman'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumLowerRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperRoman page number-->
						<xsl:when test="$pagetype='upperRoman'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumUpperRoman($count)"/>
							</pagenum>
						</xsl:when>
						<!--LowerLetter page number-->
						<xsl:when test="$pagetype='lowerLetter'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumLowerAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--UpperLetter page number-->
						<xsl:when test="$pagetype='upperLetter'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="d:PageNumUpperAlphabet($count)"/>
							</pagenum>
						</xsl:when>
						<!--Page number with dash-->
						<xsl:when test="$pagetype='numberInDash'">
							<pagenum page="special" id="{concat('page',d:GeneratePageId($myObj))}">
								<xsl:value-of select="concat('-',$count,'-')"/>
							</pagenum>
						</xsl:when>
						<!--Normal page number-->
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test="$counter=0 and d:GetSectionFront($myObj)=1">
									<pagenum page="normal" id="{concat('page',d:GeneratePageId($myObj))}">
										<xsl:value-of select="$count"/>
									</pagenum>
								</xsl:when>
								<xsl:when test="$counter=0 and d:GetSectionFront($myObj)=0">
									<pagenum page="normal" id="{concat('page',d:GeneratePageId($myObj))}">
										<xsl:value-of select="d:ReturnPageNum($myObj)"/>
									</pagenum>
								</xsl:when>
								<xsl:otherwise>
									<pagenum page="normal" id="{concat('page',d:GeneratePageId($myObj))}">
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
		<xsl:param name="Attribute" as="xs:boolean"/>
		<xsl:message terminate="no">debug in Languages</xsl:message>
		<xsl:variable name="count_lang" as="xs:integer" select="xs:integer(string-join(('0',w:r[1]/w:rPr/w:lang/count(@*)),''))"/>
		<xsl:choose>
			<!--Checking for language type CS-->
			<xsl:when test="w:r/w:rPr/w:rFonts/@w:hint='cs'">
				<xsl:choose>
					<!--Checking for bidirectional language-->
					<xsl:when test="w:r/w:rPr/w:lang/@w:bidi">
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->
							<xsl:when test="not($Attribute)">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',w:r/w:rPr/w:lang/@w:bidi,'&quot;&gt;')"/>
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
							<xsl:when test="not($Attribute)">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$doclangbidi,'&quot;&gt;')"/>
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
							<xsl:when test="not($Attribute)">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',w:r/w:rPr/w:lang/@w:eastAsia,'&quot;&gt;')"/>
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
							<xsl:when test="not($Attribute)">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$doclangeastAsia,'&quot;&gt;')"/>
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
									<xsl:when test="not($Attribute)">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',(w:r/w:rPr/w:lang/@w:val)[1],'&quot;&gt;')"/>
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
									<xsl:when test="not($Attribute)">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$doclang,'&quot;&gt;')"/>
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
									<xsl:when test="not($Attribute)">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',(w:r/w:rPr/w:lang/@w:val)[1],'&quot;&gt;')"/>
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
									<xsl:when test="not($Attribute)">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',(w:r/w:rPr/w:lang/@w:eastAsia)[1],'&quot;&gt;')"/>
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
									<xsl:when test="not($Attribute)">

										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',(w:r/w:rPr/w:lang/@w:bidi)[1],'&quot;&gt;')"/>
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
							<xsl:when test="not($Attribute)">
								<xsl:choose>
									<xsl:when test="w:r/w:rPr/w:lang/@w:val">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',(w:r/w:rPr/w:lang/@w:val)[1],'&quot;&gt;')"/>
									</xsl:when>
									<xsl:when test="w:r/w:rPr/w:lang/@w:eastAsia">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',(w:r/w:rPr/w:lang/@w:eastAsia)[1],'&quot;&gt;')"/>
									</xsl:when>
									<xsl:when test="w:r/w:rPr/w:lang/@w:bidi">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',(w:r/w:rPr/w:lang/@w:bidi)[1],'&quot;&gt;')"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of disable-output-escaping="yes" select="'&lt;p&gt;'"/>
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

	<!--Template to implement Languages on paragraph
	 NOTE NP 2022/08/26 : some errors due to wrong language count found here
	 some fixes might be required for bidi and eastAsia cases -->
	<xsl:template name="LanguagesPara">
		<xsl:param name="Attribute" as="xs:boolean"/>
		<xsl:param name="level" as="xs:string"/>

		<!--NOTE: Use w:r instead w:r[1]-->
		<xsl:variable name="count_lang" as="xs:integer" select="xs:integer(string-join(('0',count(distinct-values(w:r/w:rPr/w:lang/@w:val))),''))" />
		<xsl:choose>
			<!--Checking for language type CS-->
			<xsl:when test="w:r/w:rPr/w:rFonts/@w:hint='cs'">
				<xsl:choose>
					<!--for bidirectional language-->
					<xsl:when test="w:r/w:rPr/w:lang/@w:bidi">
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->
							<xsl:when test="not($Attribute)">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',w:r/w:rPr/w:lang/@w:bidi,'&quot;&gt;')"/>
							</xsl:when>
							<!--Creating <level> element with xml:lang attribute-->
							<xsl:otherwise>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=&quot;',w:r/w:rPr/w:lang/@w:bidi,'&quot;&gt;')"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->

							<xsl:when test="not($Attribute)">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$doclangbidi,'&quot;&gt;')"/>
							</xsl:when>
							<!--Creating <level> element with xml:lang attribute-->

							<xsl:otherwise>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=&quot;',$doclangbidi,'&quot;&gt;')"/>
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
							<xsl:when test="not($Attribute)">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',w:r/w:rPr/w:lang/@w:eastAsia,'&quot;&gt;')"/>
							</xsl:when>
							<!--Creating <level> element with xml:lang attribute-->

							<xsl:otherwise>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=&quot;',w:r/w:rPr/w:lang/@w:eastAsia,'&quot;&gt;')"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->

							<xsl:when test="not($Attribute)">
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$doclangeastAsia,'&quot;&gt;')"/>
							</xsl:when>
							<!--Creating <level> element with xml:lang attribute-->

							<xsl:otherwise>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=&quot;',$doclangeastAsia,'&quot;&gt;')"/>
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
									<xsl:when test="not($Attribute)">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',w:r/w:rPr/w:lang/@w:val,'&quot;&gt;')"/>
									</xsl:when>
									<!--Assigning language value, takes the first one by default -->
									<xsl:otherwise>
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=&quot;',distinct-values(w:r/w:rPr/w:lang/@w:val)[1],'&quot;&gt;')"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->
									<xsl:when test="not($Attribute)">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',$doclang,'&quot;&gt;')"/>
									</xsl:when>
									<!--Assingning default language value-->
									<xsl:otherwise>
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=&quot;',$doclang,'&quot;&gt;')"/>
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
									<xsl:when test="not($Attribute)">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',distinct-values(w:r/w:rPr/w:lang/@w:val)[1],'&quot;&gt;')"/>
									</xsl:when>
									<!--Assingning language value-->
									<xsl:otherwise>
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=&quot;',distinct-values(w:r/w:rPr/w:lang/@w:val)[1],'&quot;&gt;')"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="w:r/w:rPr/w:lang/@w:eastAsia">
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->
									<xsl:when test="not($Attribute)">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',w:r/w:rPr/w:lang/@w:eastAsia[1],'&quot;&gt;')"/>
									</xsl:when>
									<!--Assingning language value-->
									<xsl:otherwise>
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=&quot;',w:r/w:rPr/w:lang/@w:eastAsia[1],'&quot;&gt;')"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="w:r/w:rPr/w:lang/@w:bidi">
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->
									<xsl:when test="not($Attribute)">
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;p xml:lang=&quot;',w:r/w:rPr/w:lang/@w:bidi[1],'&quot;&gt;')"/>
									</xsl:when>
									<!--Assingning language value-->
									<xsl:otherwise>
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,' xml:lang=&quot;',w:r/w:rPr/w:lang/@w:bidi[1],'&quot;&gt;')"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:choose>
									<!--Creating <p> element with xml:lang attribute-->
									<xsl:when test="not($Attribute)">
										<xsl:value-of disable-output-escaping="yes" select="'&lt;p&gt;'"/>
									</xsl:when>
									<!--Assingning default language value-->
									<xsl:otherwise>
										<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,'&gt;')"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<!--Creating <p> element with xml:lang attribute-->
							<xsl:when test="not($Attribute)">
								<xsl:value-of disable-output-escaping="yes" select="'&lt;p&gt;'"/>
							</xsl:when>
							<!--Assingning default language value-->
							<xsl:otherwise>
								<xsl:value-of disable-output-escaping="yes" select="concat('&lt;',$level,'&gt;')"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!--Template to implement Languages-->
	<xsl:template name="PictureLanguage" as="xs:string">
		<xsl:param name="CheckLang" as="xs:string"/>
		<xsl:message terminate="no">debug in PictureLanguage</xsl:message>
		<xsl:choose>
			<!--Checking languge for picture-->
			<xsl:when test="$CheckLang='picture'">
				<xsl:variable name="count_lang" as="xs:integer" select="xs:integer(string-join(('0',../following-sibling::w:p[1]/w:r[1]/w:rPr/w:lang/count(@*)),''))"/>
				<xsl:choose>
					<!--Checking for language type eastAsia-->
					<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:rFonts/@w:hint='eastAsia'">
						<xsl:choose>
							<!--Getting value from eastasia attribute in lang tag-->
							<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia">
								<xsl:sequence select="(../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia)"/>
							</xsl:when>
							<!--Assinging default eastAsia language-->
							<xsl:otherwise>
								<xsl:sequence select="$doclangeastAsia"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<!--Checking for language type CS-->
					<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:rFonts/@w:hint='cs'">
						<xsl:choose>
							<!--Checking for bidirectional language-->
							<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi">
								<!--Getting value from bidi attribute in lang tag-->
								<xsl:sequence select="(../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi)"/>
							</xsl:when>
							<!--Assinging default bidirectional language-->
							<xsl:otherwise>
								<xsl:sequence select="$doclangbidi"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<xsl:when test="$count_lang &gt;1">
								<xsl:choose>
									<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val">
										<xsl:sequence select="(../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val)"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:sequence select="$doclang"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="$count_lang=1">
								<xsl:choose>
									<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val">
										<xsl:sequence select="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val"/>
									</xsl:when>
									<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia">
										<xsl:sequence select="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia"/>
									</xsl:when>
									<xsl:when test="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi">
										<xsl:sequence select="../following-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi"/>
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:sequence select="$doclang"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<!--Checking language for image group-->
			<xsl:when test="$CheckLang='imagegroup'">
				<xsl:variable name="count_lang" as="xs:integer" select="xs:integer(string-join(('0',../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/count(@*)),''))"/>
				<xsl:choose>
					<!--Checking for language type CS-->
					<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:rFonts/@w:hint='cs'">
						<xsl:choose>
							<!--Checking for bidirectional language-->
							<xsl:when test="(../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:bidi)">
								<!--Getting value from bidi attribute in lang tag-->
								<xsl:sequence select="(../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:bidi)"/>
							</xsl:when>
							<!--Assinging default bidirectional language-->
							<xsl:otherwise>
								<xsl:sequence select="$doclangbidi"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<!--Checking for language type eastAsia-->
					<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:rFonts/@w:hint='eastAsia'">
						<xsl:choose>
							<!--Getting value from eastasia attribute in lang tag-->
							<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:eastAsia">
								<xsl:sequence select="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:eastAsia"/>
							</xsl:when>
							<!--Assinging default eastAsia language-->
							<xsl:otherwise>
								<xsl:sequence select="$doclangeastAsia"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<xsl:when test="$count_lang &gt; 1">
								<xsl:choose>
									<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:val">
										<xsl:sequence select="(../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:val)"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:sequence select="$doclang"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="$count_lang=1">
								<xsl:choose>
									<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:val">
										<xsl:sequence select="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:val"/>
									</xsl:when>
									<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:eastAsia">
										<xsl:sequence select="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:eastAsia"/>
									</xsl:when>
									<xsl:when test="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:bidi">
										<xsl:sequence select="../../w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:rPr/w:lang/@w:bidi"/>
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:sequence select="$doclang"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<!--Checking language for table-->
			<xsl:when test="$CheckLang='Table'">
				<xsl:variable name="count_lang" as="xs:integer" select="xs:integer(string-join(('0',preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/count(@*)),''))"/>
				<xsl:choose>
					<!--Checking for language type eastAsia-->
					<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:rFonts/@w:hint='eastAsia'">
						<xsl:choose>
							<!--Getting value from eastasia attribute in lang tag-->

							<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia">
								<xsl:sequence select="(preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia)"/>
							</xsl:when>
							<!--Assinging default eastAsia language-->

							<xsl:otherwise>
								<xsl:sequence select="$doclangeastAsia"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<!--Checking for language type CS-->
					<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:rFonts/@w:hint='cs'">
						<xsl:choose>
							<!--Checking for bidirectional language-->
							<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi">
								<!--Getting value from bidi attribute in lang tag-->
								<xsl:sequence select="(preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi)"/>
							</xsl:when>
							<!--Assinging default bidirectional language-->
							<xsl:otherwise>
								<xsl:sequence select="$doclangbidi"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<xsl:when test="$count_lang &gt; 1">
								<xsl:choose>
									<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val">
										<xsl:sequence select="(preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val)"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:sequence select="$doclang"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="$count_lang = 1">
								<xsl:choose>
									<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val">
										<xsl:sequence select="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:val"/>
									</xsl:when>
									<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia">
										<xsl:sequence select="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:eastAsia"/>
									</xsl:when>
									<xsl:when test="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi">
										<xsl:sequence select="preceding-sibling::w:p[1]/w:r/w:rPr/w:lang/@w:bidi"/>
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:sequence select="$doclang"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
		</xsl:choose>
	</xsl:template>

	<xsl:template name="GetBdoLanguages">
		<xsl:param name="runner"/>
        <xsl:message terminate="no">debug in GetBdoLanguages</xsl:message>
        <xsl:choose>
			<!-- Complex Script Font -->
            <xsl:when test="$runner/w:rPr/w:rFonts/@w:hint='cs'">
                <xsl:choose>
                    <xsl:when test="$runner/w:rPr/w:lang/@w:bidi">
                        <xsl:value-of select="$runner/w:rPr/w:lang/@w:bidi"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="$doclangbidi"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
			<!-- East Asia Font -->
            <xsl:when test="$runner/w:rPr/w:rFonts/@w:hint='eastAsia'">
                <xsl:choose>
                    <xsl:when test="$runner/w:rPr/w:lang/@w:eastAsia">
                        <xsl:value-of select="$runner/w:rPr/w:lang/@w:eastAsia"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="$doclangeastAsia"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
			<!-- Default Font -->
            <xsl:otherwise>
				<xsl:choose>
					<!-- A bidirectionnal lang is set -->
					<xsl:when test="$runner/w:rPr/w:lang/@w:bidi">
						<xsl:value-of select="$runner/w:rPr/w:lang/@w:bidi"/>
					</xsl:when>
					<!-- An east asia lang is set -->
					<xsl:when test="$runner/w:rPr/w:lang/@w:eastAsia">
						<xsl:value-of select="$runner/w:rPr/w:lang/@w:eastAsia"/>
					</xsl:when>
					<!-- An alternative lang (but no bidi or east asia) is set -->
					<xsl:when test="$runner/w:rPr/w:lang/@w:val">
						<xsl:value-of select="$runner/w:rPr/w:lang/@w:val"/>
					</xsl:when>
					<!-- No lang set, return default doc lang -->
					<xsl:otherwise>
						<xsl:value-of select="$doclang"/>
					</xsl:otherwise>
				</xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

	<xsl:template name="TempCharacterStyle">
		<xsl:param name ="characterStyle" as="xs:boolean"/>
		<xsl:choose>
			<xsl:when test="$characterStyle">
				<xsl:choose>
					<xsl:when test="../w:pPr/w:ind[@w:left] and ../w:pPr/w:ind[@w:right] and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:caps and w:rPr/w:color and w:t">
						<xsl:variable name="val" as="xs:integer" select="../w:pPr/w:ind/@w:left"/>
						<xsl:variable name="val_left" as="xs:integer" select="($val div 1440)"/>
						<xsl:variable name="valright" as="xs:integer" select="../w:pPr/w:ind/@w:right"/>
						<xsl:variable name="val_right" as="xs:integer" select="($valright div 1440)"/>
						<xsl:variable name="val_color" as="xs:string" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';text-transform:uppercase',';text-indent:','right=',$val_right,'in',';left=',$val_left,'in')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="../w:pPr/w:ind[@w:left] and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:caps and w:rPr/w:color and w:t">
						<xsl:variable name="val" as="xs:integer" select="../w:pPr/w:ind/@w:left"/>
						<xsl:variable name="val_left" as="xs:integer" select="($val div 1440)"/>
						<xsl:variable name="val_color" as="xs:string" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';text-transform:uppercase',';text-indent:',$val_left,'in')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="../w:pPr/w:ind[@w:right] and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:caps and w:rPr/w:color and w:t">
						<xsl:variable name="val" as="xs:integer" select="../w:pPr/w:ind/@w:right"/>
						<xsl:variable name="val_right" as="xs:integer" select="($val div 1440)"/>
						<xsl:variable name="val_color" as="xs:string" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';text-transform:uppercase',';text-indent:',$val_right,'in')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="../w:pPr/w:jc and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:caps and w:rPr/w:color and w:t">
						<xsl:variable name="val" as="xs:string" select="../w:pPr/w:jc/@w:val"/>
						<xsl:variable name="val_color" as="xs:string" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';text-transform:uppercase',';text-align:',$val)}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="w:rPr/w:u and w:rPr/w:strike and w:rPr/w:caps and w:rPr/w:color and w:t">
						<xsl:variable name="val_color" as="xs:string" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';text-transform:uppercase')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>


					<xsl:when test="../w:pPr/w:ind[@w:left] and ../w:pPr/w:ind[@w:right] and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:smallCaps and w:rPr/w:color and w:t">
						<xsl:variable name="val" as="xs:integer" select="../w:pPr/w:ind/@w:left"/>
						<xsl:variable name="val_left" as="xs:integer" select="($val div 1440)"/>
						<xsl:variable name="valright" as="xs:integer" select="../w:pPr/w:ind/@w:right"/>
						<xsl:variable name="val_right" as="xs:integer" select="($valright div 1440)"/>
						<xsl:variable name="val_color" as="xs:string" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';font-variant:small-caps',';text-indent:','right=',$val_right,'in',';left=',$val_left,'in')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="../w:pPr/w:ind[@w:left] and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:smallCaps and w:rPr/w:color and w:t">
						<xsl:variable name="val" as="xs:integer" select="../w:pPr/w:ind/@w:left"/>
						<xsl:variable name="val_left" as="xs:integer" select="($val div 1440)"/>
						<xsl:variable name="val_color" as="xs:string" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';font-variant:small-caps',';text-indent:',$val_left,'in')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="../w:pPr/w:ind[@w:right] and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:smallCaps and w:rPr/w:color and w:t">
						<xsl:variable name="val" as="xs:integer" select="../w:pPr/w:ind/@w:right"/>
						<xsl:variable name="val_right" as="xs:integer" select="($val div 1440)"/>
						<xsl:variable name="val_color" as="xs:string" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';font-variant:small-caps',';text-indent:',$val_right,'in')}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="../w:pPr/w:jc and w:rPr/w:u and w:rPr/w:strike and w:rPr/w:smallCaps and w:rPr/w:color and w:t">
						<xsl:variable name="val" as="xs:string" select="../w:pPr/w:jc/@w:val"/>
						<xsl:variable name="val_color" as="xs:string" select="w:rPr/w:color/@w:val"/>
						<span class="{concat('text:Underline line-through;color:#',$val_color,';font-variant:small-caps',';text-align:',$val)}">
							<xsl:value-of select="w:t"/>
						</span>
					</xsl:when>

					<xsl:when test="w:rPr/w:u and w:rPr/w:strike and w:rPr/w:smallCaps and w:rPr/w:color and w:t">
						<xsl:variable name="val_color" as="xs:string" select="w:rPr/w:color/@w:val"/>
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
