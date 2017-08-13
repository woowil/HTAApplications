<?xml version="1.0" encoding="ISO-8859-1" standalone="no"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format">

<!-- http://www.w3.org/TR/xslt -->

<xsl:import href="mba-xsl-common.xsl"/>

<xsl:template match="mba">

<html dir='ltr'><head><title>Platform Systems Network</title></head>
<body style="{$BODY}" scroll="auto"><br/>
<table cellspacing="0" cellpadding="0" align="center" border="0" style="{$TABLE}">
<thead>
<tr>
	<td colspan="27">
	<table cellspacing="0" width="100%" cellpadding="0" align="center" border="0">
	<tr>		
		<th style="{$TH}" width="10%" nowrap="true">Updated date</th><td style="{$TD}" width="10%" nowrap="true"><xsl:value-of select="//mba/mba_info/updated_date"/></td>
		<th style="{$TH}" width="10%" nowrap="true">Updated by</th><td style="{$TD}"><xsl:value-of select="//mba/mba_info/updated_by"/></td>		
	</tr>
	<tr>
		<th style="{$TH}">Information</th><td style="{$TD}" colspan="3"><xsl:value-of select="//mba/mba_info/information"/></td>
	</tr>
	</table>
	<div style="color:silver;font-size:9px;font-style:italic;margin-left:100px"><br/><xsl:value-of select="//mba/mba_info/powered_by"/></div><br/>
	</td>
</tr>
<tr>
	<th style="{$TH}"><xsl:text> </xsl:text></th>
<xsl:for-each select="//mba_header/head">
	<th style="{$TH}"><xsl:value-of select="."/></th>
</xsl:for-each>
	<th style="{$TH}"><xsl:text> </xsl:text></th>
</tr>
</thead>
<tbody>
<xsl:for-each select="//mba/mba_config">
<xsl:sort select="name" data-type="text" order="ascending" lang="en"/>
	<tr>
	<td style="{$TD}"><xsl:number value="position()" format="001"/></td>
<xsl:for-each select="*">
	<td style="{$TD}" nowrap="true"><xsl:value-of select="."/> </td>
</xsl:for-each>
	<td style="{$TD}"><xsl:number value="position()" format="001"/></td>
</tr>
</xsl:for-each>


</tbody>
<tfoot>
<tr>
	<th style="{$TH}"><xsl:text> </xsl:text></th>
<xsl:for-each select="//mba_header/head">
	<th style="{$TH}"><xsl:value-of select="."/></th>
	</xsl:for-each>
<th style="{$TH}"><xsl:text> </xsl:text></th>
</tr>
</tfoot>
</table>
</body>
</html>

</xsl:template>

</xsl:stylesheet>
