<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" version="1.0"
                xmlns:px="http://www.daisy.org/ns/pipeline/xproc"
                xmlns:cx="http://xmlcalabash.com/ns/extensions"
                type="px:oox2Daisy" name="main">

	<p:option name="source" required="true"/>
	<p:option name="output-dir" required="true"/>
	<p:option name="Title" select="''"/>
	<p:option name="Creator" select="''"/>
	<p:option name="Publisher" select="''"/>
	<p:option name="UID" select="''"/>
	<p:option name="Subject" select="''"/>
	<p:option name="prmTRACK" select="'NoTrack'"/>
	<p:option name="Version" select="'14'"/>
	<p:option name="Custom" select="''"/>
	<p:option name="MasterSub" select="false()"/>
	<p:option name="ImageSizeOption" select="'original'"/>
	<p:option name="DPI" select="96"/>
	<p:option name="CharacterStyles" select="false()"/>
	<p:option name="MathML" select="map{'wdTextFrameStory':[],
	                                    'wdFootnotesStory':[],
	                                    'wdMainTextStory':[]}"/>
	<p:output port="result"/>

	<p:xslt template-name="main" cx:serialize="true">
		<p:input port="source">
			<p:empty/>
		</p:input>
		<p:input port="stylesheet">
			<p:document href="oox2Daisy.xsl"/>
		</p:input>
		<p:with-param name="OriginalInputFile" select="$source"/>
		<p:with-param name="InputFile" select="$source"/>
		<p:with-param name="OutputDir" select="$output-dir"/>
		<p:with-param name="FinalOutputDir" select="$output-dir"/>
		<p:with-param name="Title" select="$Title"/>
		<p:with-param name="Creator" select="$Creator"/>
		<p:with-param name="Publisher" select="$Publisher"/>
		<p:with-param name="UID" select="$UID"/>
		<p:with-param name="Subject" select="$Subject"/>
		<p:with-param name="prmTRACK" select="$prmTRACK"/>
		<p:with-param name="Version" select="$Version"/>
		<p:with-param name="Custom" select="$Custom"/>
		<p:with-param name="MasterSub" select="$MasterSub"/>
		<p:with-param name="ImageSizeOption" select="$ImageSizeOption"/>
		<p:with-param name="DPI" select="$DPI"/>
		<p:with-param name="CharacterStyles" select="$CharacterStyles"/>
		<p:with-param name="MathML" select="$MathML"/>
	</p:xslt>
	<p:group>
		<p:variable name="output-file"
		            select="concat($output-dir,
		                           replace(replace($source,'^.*/([^/]*?)(\.[^/\.]*)?$','$1.xml'),',','_'))"/>
		<p:store name="store-xml">
			<p:with-option name="href" select="$output-file"/>
		</p:store>
		<p:store name="store-css">
			<p:input port="source">
				<p:inline><css/></p:inline>
			</p:input>
			<p:with-option name="href" select="concat($output-dir,'dtbookbasic.css')"/>
		</p:store>
		<p:load cx:depends-on="store-xml">
			<p:with-option name="href" select="$output-file"/>
		</p:load>
		<p:identity cx:depends-on="store-css"/>
	</p:group>

</p:declare-step>
