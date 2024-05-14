<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" version="1.0"
                xmlns:px="http://www.daisy.org/ns/pipeline/xproc"
                xmlns:cx="http://xmlcalabash.com/ns/extensions"
				xmlns:xs="http://www.w3.org/2001/XMLSchema"
                type="px:word-to-dtbook.script" name="main">

	<p:documentation xmlns="http://www.w3.org/1999/xhtml">
		<h1 px:role="name">Word to dtbook</h1>
		<p px:role="desc" xml:space="preserve">Transforms a Microsoft Office Word (.docx) document into a DTBook XML file.</p>
		<a px:role="homepage" href="http://daisy.github.io/pipeline/Get-Help/User-Guide/Scripts/word-to-dtbook/">
			Online documentation
		</a>
		<!--<address>
			Authors:
			<dl px:role="author">
				<dt>Name:</dt>
				<dd px:role="name">Bert Frees</dd>
				<dt>E-mail:</dt>
				<dd><a px:role="contact" href="mailto:bertfrees@gmail.com">bertfrees@gmail.com</a></dd>
				<dt>Organization:</dt>
				<dd px:role="organization" href="http://www.sbs-online.ch/">SBS</dd>
			</dl>
			<dl px:role="author">
				<dt>Name:</dt>
				<dd px:role="name">Jostein Austvik Jacobsen</dd>
				<dt>E-mail:</dt>
				<dd><a px:role="contact" href="mailto:josteinaj@gmail.com">josteinaj@gmail.com</a></dd>
				<dt>Organization:</dt>
				<dd px:role="organization" href="http://www.nlb.no/">NLB</dd>
			</dl>
		</address>-->
	</p:documentation>

	<p:option name="source" required="true" px:type="anyFileURI">
		<p:documentation>
			<h2 px:role="name">Input Docx file</h2>
			<p px:role="desc" xml:space="preserve">The document you want to convert.</p>
		</p:documentation>
	</p:option>
	<p:option name="document-output-dir" required="true"/>
	<p:option name="document-output-file" select="concat(
		$document-output-dir,
		replace(replace($source,'^.*/([^/]*?)(\.[^/\.]*)?$','$1.xml'),',','_')
	)"/>
	<p:option name="document-output-test" select="concat(
		$document-output-dir,
		replace(replace($source,'^.*/([^/]*?)(\.[^/\.]*)?$','$1_test.xml'),',','_')
	)"/>
	<p:option name="tempSource" select="$source"/>
	<p:option name="pipeline-output-dir" select="$document-output-dir"/>
	<p:option name="Title" select="''">
		<p:documentation xmlns="http://www.w3.org/1999/xhtml">
			<h2 px:role="name">Document title</h2>
			<p px:role="desc"></p>
		</p:documentation>
	</p:option>
	<p:option name="Creator" select="''">
		<p:documentation xmlns="http://www.w3.org/1999/xhtml">
			<h2 px:role="name">Document author</h2>
			<p px:role="desc"></p>
		</p:documentation>
	</p:option>
	<p:option name="Publisher" select="''">
		<p:documentation xmlns="http://www.w3.org/1999/xhtml">
			<h2 px:role="name">Document publisher</h2>
			<p px:role="desc"></p>
		</p:documentation>
	</p:option>
	<p:option name="UID" select="''"/>
	<p:option name="Subject" select="''"/>
	<p:option name="acceptRevisions" select="true()"/>
	<p:option name="Version" select="'14'"/>

	<!-- discarding math type equations preprocessing
	<p:option name="MathML" select="map{'wdTextFrameStory':[],
	                                    'wdFootnotesStory':[],
	                                    'wdMainTextStory':[]
	                                    }" />-->
	<!-- cx:as="map(xs:string,xs:string*)" -->
	<p:option name="MasterSub" select="false()" cx:as="xs:boolean" />
	<!-- from settings  -->
	<p:option name="Custom" select="''"/>
	<p:option name="ImageSizeOption" select="'original'">
		<p:documentation xmlns="http://www.w3.org/1999/xhtml">
			<h2 px:role="name">Image size option</h2>
			<p px:role="desc"></p>
		</p:documentation>
		<p:pipeinfo>
			<px:type>
				<choice xmlns:a="http://relaxng.org/ns/compatibility/annotations/1.0">
					<value>original</value>
					<a:documentation xml:lang="en">Keep image size</a:documentation>
					<value>resize</value>
					<a:documentation xml:lang="en">Resize images</a:documentation>
					<value>resample</value>
					<a:documentation xml:lang="en">Resample images</a:documentation>
				</choice>
			</px:type>
		</p:pipeinfo>
	</p:option>
	<p:option name="DPI" select="96" cx:as="xs:integer">
		<p:documentation xmlns="http://www.w3.org/1999/xhtml">
			<h2 px:role="name">Image resampling value</h2>
			<p px:role="desc">Image resampling targeted resolution in dpi (dot-per-inch)</p>
		</p:documentation>
	</p:option>
	<p:option name="CharacterStyles" select="false()" cx:as="xs:boolean">
		<p:documentation xmlns="http://www.w3.org/1999/xhtml">
			<h2 px:role="name">Translate character styles</h2>
			<p px:role="desc"></p>
		</p:documentation>
	</p:option>
	<p:option name="FootnotesPosition" select="'end'">
		<p:documentation xmlns="http://www.w3.org/1999/xhtml">
			<h2 px:role="name">Footnotes position</h2>
			<p px:role="desc">Position of footnotes in content</p>
		</p:documentation>
		<p:pipeinfo>
			<px:type>
				<choice xmlns:a="http://relaxng.org/ns/compatibility/annotations/1.0">
					<value>inline</value>
					<a:documentation xml:lang="en">Inline note in content (after the paragraph containing its first reference)</a:documentation>
					<value>end</value>
					<a:documentation xml:lang="en">Put notes at the end of a level defined in footnotes insertion level</a:documentation>
					<value>page</value>
					<a:documentation xml:lang="en">Put the notes near the page break</a:documentation>
				</choice>
			</px:type>
		</p:pipeinfo>
	</p:option>
	<p:option name="FootnotesLevel" select="0" cx:as="xs:integer">
		<p:documentation xmlns="http://www.w3.org/1999/xhtml">
			<h2 px:role="name">Footnotes insertion level</h2>
			<p px:role="desc">Closest level into which notes are inserted in content.</p>
		</p:documentation>
	</p:option>
	<p:option name="FootnotesNumbering" cx:as="xs:string" select="'none'">
		<p:documentation xmlns="http://www.w3.org/1999/xhtml">
			<h2 px:role="name">Footnotes position</h2>
			<p px:role="desc">Position of footnotes in content</p>
		</p:documentation>
		<p:pipeinfo>
			<px:type>
				<choice xmlns:a="http://relaxng.org/ns/compatibility/annotations/1.0">
					<value>none</value>
					<a:documentation xml:lang="en">Disable note numbering in output</a:documentation>
					<value>word</value>
					<a:documentation xml:lang="en">Use original word numbering</a:documentation>
					<value>number</value>
					<a:documentation xml:lang="en">Use custom numbering, starting from the footnotes start value</a:documentation>
				</choice>
			</px:type>
		</p:pipeinfo>
	</p:option>
	<p:option name="FootnotesStartValue" cx:as="xs:integer" select="1">
		<p:documentation xmlns="http://www.w3.org/1999/xhtml">
			<h2 px:role="name">Footnotes starting value</h2>
			<p px:role="desc">If footnotes numbering is required, start the notes numbering process from this value</p>
		</p:documentation>
	</p:option>
	<p:option name="FootnotesNumberingPrefix" cx:as="xs:string?" select="''">
		<p:documentation xmlns="http://www.w3.org/1999/xhtml">
			<h2 px:role="name">Footnotes number prefix</h2>
			<p px:role="desc">Add an prefix before the note numbering</p>
		</p:documentation>
	</p:option>
	<p:option name="FootnotesNumberingSuffix" cx:as="xs:string?" select="''"/>
		<p:documentation xmlns="http://www.w3.org/1999/xhtml">
			<h2 px:role="name">Footnotes number suffix</h2>
			<p px:role="desc">Add a text between the notes number and its textual content.</p>
		</p:documentation>
	<p:output port="result" sequence="true"/>

	<p:xslt template-name="main" name="convert-to-dtbook" cx:serialize="true">
		<p:input port="source">
			<p:empty/>
		</p:input>
		<p:input port="stylesheet">
			<p:document href="oox2Daisy.xsl"/>
		</p:input>
		<p:with-param name="OriginalInputFile" select="$source"/>
		<p:with-param name="InputFile" select="$tempSource"/>
		<p:with-param name="OutputDir" select="$document-output-dir"/>
		<p:with-param name="FinalOutputDir" select="$pipeline-output-dir"/>
		<p:with-param name="Title" select="$Title"/>
		<p:with-param name="Creator" select="$Creator"/>
		<p:with-param name="Publisher" select="$Publisher"/>
		<p:with-param name="UID" select="$UID"/>
		<p:with-param name="Subject" select="$Subject"/>
		<p:with-param name="acceptRevisions" select="$acceptRevisions"/>
		<p:with-param name="Version" select="$Version"/>
		<p:with-param name="Custom" select="$Custom"/>
		<p:with-param name="MasterSub" select="$MasterSub"/>
		<p:with-param name="ImageSizeOption" select="$ImageSizeOption"/>
		<p:with-param name="DPI" select="$DPI"/>
		<p:with-param name="CharacterStyles" select="$CharacterStyles"/>
		<p:with-param name="FootnotesPosition" select="$FootnotesPosition"/>
		<p:with-param name="FootnotesLevel" select="$FootnotesLevel"/>
		<p:with-param name="FootnotesNumbering" select="$FootnotesNumbering"/>
		<p:with-param name="FootnotesStartValue" select="$FootnotesStartValue"/>
		<p:with-param name="FootnotesNumberingPrefix" select="$FootnotesNumberingPrefix"/>
		<p:with-param name="FootnotesNumberingSuffix" select="$FootnotesNumberingSuffix"/>
	</p:xslt>
	<p:store name="store-xml">
		<p:with-option name="href" select="$document-output-file"/>
	</p:store>
	<p:load cx:depends-on="store-xml">
		<p:with-option name="href" select="$document-output-file"/>
	</p:load>


</p:declare-step>
