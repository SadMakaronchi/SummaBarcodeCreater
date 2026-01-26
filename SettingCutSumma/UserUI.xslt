<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:frmwrk="Corel Framework Data" exclude-result-prefixes="frmwrk">
	<xsl:output method="xml" encoding="UTF-8" indent="yes"/>

	<frmwrk:uiconfig>
		<frmwrk:compositeNode xPath="/uiConfig/commandBars/commandBarData[@guid='f3016f3c-2847-4557-b61a-a2d05319cf18']"/>
		<frmwrk:compositeNode xPath="/uiConfig/frame"/>
	</frmwrk:uiconfig>

	<xsl:template match="node()|@*">
		<xsl:copy>
			<xsl:apply-templates select="node()|@*"/>
		</xsl:copy>
	</xsl:template>

	<xsl:template match="uiConfig/commandBars/commandBarData[@guid='f3016f3c-2847-4557-b61a-a2d05319cf18']/menubar/modeData[@guid='76d73481-9076-44c9-821c-52de9408cce2']/item[@guidRef='6c91d5ab-d648-4364-96fb-3e71bcfaf70d']">
		<xsl:copy-of select="."/>
		<item guidRef="8ab44227-8f35-47f3-8e6c-0fc096571ac8"/>
	</xsl:template>
</xsl:stylesheet>