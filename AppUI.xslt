<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:frmwrk="Corel Framework Data">
	<xsl:output method="xml" encoding="UTF-8" indent="yes"/>
	<frmwrk:uiconfig>
		<frmwrk:applicationInfo userConfiguration="true" />
	</frmwrk:uiconfig>
	<!--Копирование всего интерфейса-->
	<xsl:template match="node()|@*">
		<xsl:copy>
			<xsl:apply-templates select="node()|@*"/>
		</xsl:copy>
	</xsl:template>
	<!--Определение пункта меню-->
	<xsl:template match="uiConfig/items">
		<xsl:copy>
			<xsl:apply-templates select="node()|@*"/>
			<itemData guid="8ab44227-8f35-47f3-8e6c-0fc096571ac8"
					  noBmpOnMenu="true"
					  flyoutBarRef="d0b1b08c-1e1e-4793-9f44-acdab3319a18"
					  type="flyout"
					  userCreated="true"
					  userToolTip="Меню плагина для Summa  by sad_makaronchi"
					  userCaption="Меню плагина для Summa  by sad_makaronchi"/>
			<itemData guid="41a13089-1d92-4b2d-a224-f11daefb7b83"
					  noBmpOnMenu="true"
					  type="checkButton"
					  check="*Toolbar('109fb38b-2059-4b16-8e3c-64a2ac379687')"
					  dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
					  userCaption="Панель добавления меток на макет"
                      enable="true"/>
			<itemData guid="90d2a1b9-6465-4750-899c-92de213b05d2"
					  type="wpfhost"
				      hostedType="Addons\SummaBarcodeCreater\SummaBarcodeCreater.dll,SummaBarcodeCreater.ControlUI"
					  enable="true"/>
			<itemData guid="8a42b4c7-3a79-4199-9f00-5c17d86ab719"
			          type="wpfhost"
			          hostedType="Addons\SummaBarcodeCreater\SummaBarcodeCreater.dll,SummaBarcodeCreater.ControlUI2"
			          enable="true"/>

		</xsl:copy>
	</xsl:template>
	<!--Определение самого меню-->
	<xsl:template match="uiConfig/commandBars">
		<xsl:copy>
			<xsl:apply-templates select="node()|@*"/>
			<commandBarData guid="109fb38b-2059-4b16-8e3c-64a2ac379687"
							nonLocalizableName="Панель добавления меток с баркодом для SummaCut"
							userCaption="Панель добавления меток с баркодом для SummaCut">
				<toolbar dock="fill">
					<item dock="fill" margin="0,0,0,0" guidRef="90d2a1b9-6465-4750-899c-92de213b05d2"/>
					<item dock="fill" margin="0,0,0,0" guidRef="8a42b4c7-3a79-4199-9f00-5c17d86ab719"/>
				</toolbar>
			</commandBarData>

			<commandBarData guid="d0b1b08c-1e1e-4793-9f44-acdab3319a18"
							type="menu"
							nonLocalizableName="Меню плагина для Summa  by sad_makaronchi"
							flyout="true">
				<menu>
					<item guidRef="41a13089-1d92-4b2d-a224-f11daefb7b83"/>

				</menu>
			</commandBarData>
		</xsl:copy>
	</xsl:template>



</xsl:stylesheet>


