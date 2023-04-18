<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" version="1.0"
                xmlns:px="http://www.daisy.org/ns/pipeline/xproc"
                xmlns:cx="http://xmlcalabash.com/ns/extensions"
                xmlns:xs="http://www.w3.org/2001/XMLSchema"
                type="px:dtbook-narrator" name="main">

    <p:input port="source" primary="true" sequence="true" />
    <p:output port="result" primary="true" sequence="true" />

    <p:option name="publisher" required="false" select="''" />

    <p:xslt px:message="narrator-metadata">
        <p:input port="stylesheet"><p:document href="xsl/narrator-metadata.xsl"/></p:input>
        <p:with-param name="publisherValue" select="$publisher"/>
    </p:xslt>
    <p:xslt px:message="narrator-headings-r14">
        <p:input port="stylesheet"><p:document href="xsl/narrator-headings-r14.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt px:message="narrator-headings-r100">
        <p:input port="stylesheet"><p:document href="xsl/narrator-headings-r100.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt px:message="narrator-title">
        <p:input port="stylesheet"><p:document href="xsl/narrator-title.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt px:message="narrator-lists">
        <p:input port="stylesheet"><p:document href="xsl/narrator-lists.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>

</p:declare-step>
