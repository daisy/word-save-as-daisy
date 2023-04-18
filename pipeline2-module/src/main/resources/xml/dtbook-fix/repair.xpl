<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" version="1.0"
                xmlns:px="http://www.daisy.org/ns/pipeline/xproc"
                xmlns:cx="http://xmlcalabash.com/ns/extensions"
                xmlns:xs="http://www.w3.org/2001/XMLSchema"
                type="px:dtbook-repair" name="main">

    <p:input port="source" primary="true" sequence="true" />
    <p:output port="result" primary="true" sequence="true" />

    <p:option name="fixCharset" required="false" select="false()"/>
    <p:choose>
        <p:when test="$fixCharset">
            <!--TODO : port the pipeline1 se_tpb_dtbookFix "CharsetExecutor" as a step and call it here -->
            <p:identity />
        </p:when>
        <p:otherwise><p:identity /></p:otherwise>
    </p:choose>
    <p:xslt px:message="repair-levelnormalizer">
        <p:input port="stylesheet"><p:document href="xsl/repair-levelnormalizer.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt px:message="repair-levelsplitter">
        <p:input port="stylesheet"><p:document href="xsl/repair-levelsplitter.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt px:message="repair-add-levels">
        <p:input port="stylesheet"><p:document href="xsl/repair-add-levels.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt px:message="repair-remove-illegal-headings">
        <p:input port="stylesheet"><p:document href="xsl/repair-remove-illegal-headings.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt px:message="repair-flatten-redundant-nesting">
        <p:input port="stylesheet"><p:document href="xsl/repair-flatten-redundant-nesting.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt px:message="repair-complete-structure">
        <p:input port="stylesheet"><p:document href="xsl/repair-complete-structure.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt px:message="repair-lists">
        <p:input port="stylesheet"><p:document href="xsl/repair-lists.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt px:message="repair-idref">
        <p:input port="stylesheet"><p:document href="xsl/repair-idref.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt px:message="repair-remove-empty-elements">
        <p:input port="stylesheet"><p:document href="xsl/repair-remove-empty-elements.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt px:message="repair-pagenum-type">
        <p:input port="stylesheet"><p:document href="xsl/repair-pagenum-type.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt px:message="repair-metadata">
        <p:input port="stylesheet"><p:document href="xsl/repair-metadata.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>

    <!--TODO : port the pipeline1 se_tpb_dtbookFix "InvalidURIExecutor" as a step and call it here -->

</p:declare-step>
