<xsl:stylesheet version="3.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xpath-default-namespace="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:my="http://localhost"
  exclude-result-prefixes="xs my"
  expand-text="yes">

  <xsl:variable name="input-sheets" select="collection('?select=*.xml')"/>

  <xsl:variable name="master-column-keys" as="xs:string+">
    <xsl:for-each-group select="$input-sheets/Workbook/Worksheet/Table/Row[1]/Cell"
                        group-by="substring-before(my:column-key(.), '$')">
      <xsl:for-each select="distinct-values(current-group()/my:column-key(.))">
        <xsl:sort select="substring-after(., '$')" data-type="number" order="ascending"/>
        <xsl:sequence select="."/>
      </xsl:for-each>
    </xsl:for-each-group>
  </xsl:variable>

  <xsl:template mode="wrapper #default" match="@* | node()">
    <xsl:copy>
      <xsl:apply-templates mode="#current" select="@* | node()"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="/">
    <!-- For debugging
    <xsl:value-of select="$master-column-keys" separator="&#xA;"/>
    -->
    <xsl:apply-templates mode="wrapper" select="$input-sheets[1]"/>
  </xsl:template>

  <xsl:template mode="wrapper" match="Worksheet">
    <Worksheet ss:Name="merged-spreadsheet">
      <xsl:apply-templates mode="#current"/>
    </Worksheet>
  </xsl:template>

  <xsl:template mode="wrapper" match="Table">
    <Table>
      <xsl:apply-templates mode="#current" select="@*"/>
      <Row>
        <xsl:for-each select="$master-column-keys">
          <Cell>
            <Data ss:Type="String">{substring-before(.,'$')}</Data>
          </Cell>
        </xsl:for-each>
      </Row>
      <xsl:apply-templates select="$input-sheets/Workbook/Worksheet/Table/Row[position() gt 1]"/>
    </Table>
  </xsl:template>

  <xsl:template mode="wrapper" match="Table/@ss:ExpandedColumnCount | Table/@ss:ExpandedRowCount"/>

  <xsl:template match="Row">
    <xsl:copy>
      <xsl:apply-templates select="@*"/>
      <xsl:apply-templates select="Cell">
        <xsl:sort select="my:master-position-of(.)"/>
      </xsl:apply-templates>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="Cell">
    <xsl:copy>
      <xsl:apply-templates select="@*"/>
      <xsl:attribute name="ss:Index" select="my:master-position-of(.)"/>
      <xsl:apply-templates/>
    </xsl:copy>
  </xsl:template>

  <xsl:function name="my:master-position-of" as="xs:integer">
    <xsl:param name="cell" as="element(Cell)"/>
    <xsl:sequence select="index-of($master-column-keys, my:column-key(my:column-heading($cell)))"/>
  </xsl:function>

  <xsl:function name="my:column-heading" as="element(Cell)">
    <xsl:param name="cell" as="element(Cell)"/>
    <xsl:variable name="previous-reset-cell" select="$cell/preceding-sibling::Cell[@ss:Index][1]"/>
    <xsl:variable name="column-position" select="if ($cell/@ss:Index)           then number($cell/@ss:Index)
                                            else if (not($previous-reset-cell)) then my:position($cell)
                                            else number($previous-reset-cell/@ss:Index)
                                                 + my:position($cell)
                                                 - my:position($previous-reset-cell)"/>
    <xsl:sequence select="root($cell)/Workbook/Worksheet/Table/Row[1]/Cell[position() eq $column-position]"/>
  </xsl:function>

  <xsl:function name="my:position" as="xs:integer">
    <xsl:param name="cell" as="element(Cell)"/>
    <xsl:sequence select="1 + count($cell/preceding-sibling::Cell)"/>
  </xsl:function>

  <xsl:function name="my:column-key" as="xs:string">
    <xsl:param name="heading" as="element(Cell)"/>
    <xsl:sequence select="concat($heading/Data,
                                 '$',
                                 1 + count($heading/preceding-sibling::Cell[Data eq $heading/Data])
                                )"/>
  </xsl:function>

</xsl:stylesheet>
