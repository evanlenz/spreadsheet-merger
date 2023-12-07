<xsl:stylesheet version="3.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xpath-default-namespace="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:my="http://localhost"
  exclude-result-prefixes="xs my"
  expand-text="yes">

  <!-- Find all .xml files in the current directory -->
  <!-- ASSUMPTION: They are all in the "XML Spreadsheet 2003" format -->
  <xsl:variable name="input-sheets" select="collection('?select=*.xml')"/>

  <!-- Get a master list of column headings for the result, aggregated from all the given spreadsheets -->
  <xsl:variable name="master-column-keys" as="xs:string+">
    <xsl:for-each-group select="$input-sheets/Workbook/Worksheet/Table/Row[1]/Cell"
                        group-by="substring-before(my:column-key(.), $column-key-separator)">
      <xsl:for-each select="distinct-values(current-group()/my:column-key(.))">
        <xsl:sort select="substring-after(., $column-key-separator)" data-type="number" order="ascending"/>
        <xsl:sequence select="."/>
      </xsl:for-each>
    </xsl:for-each-group>
  </xsl:variable>

  <!-- By default, copy everything unchanged, both for the wrapper and the
       default mode, which is used to process the actual data -->
  <xsl:template mode="wrapper #default" match="@* | node()">
    <xsl:copy>
      <xsl:apply-templates mode="#current" select="@* | node()"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="/">
    <!-- For debugging
    <xsl:value-of select="$master-column-keys" separator="&#xA;"/>
    -->
    <!-- Use the outer wrapping of the first spreadsheet as the basis for creating the merged spreadsheet -->
    <xsl:apply-templates mode="wrapper" select="$input-sheets[1]"/>
  </xsl:template>

  <!-- Create one worksheet with the merged data -->
  <xsl:template mode="wrapper" match="Worksheet">
    <Worksheet ss:Name="merged-spreadsheet">
      <xsl:apply-templates mode="#current"/>
    </Worksheet>
  </xsl:template>

  <!-- Generate the final table columns and all rows from all spreadsheets -->
  <xsl:template mode="wrapper" match="Table">
    <Table>
      <xsl:apply-templates mode="#current" select="@*"/>
      <Row>
        <!-- One column for each of the column keys we arrived at above -->
        <xsl:for-each select="$master-column-keys">
          <Cell>
            <Data ss:Type="String">{substring-before(., $column-key-separator)}</Data>
          </Cell>
        </xsl:for-each>
      </Row>
      <!-- Now go process all the data rows in all the spreadsheets -->
      <xsl:apply-templates select="$input-sheets/Workbook/Worksheet/Table/Row[position() gt 1]"/>
    </Table>
  </xsl:template>

  <!-- Strip out these attributes which aren't required and won't be accurate for the merged spreadsheet -->
  <xsl:template mode="wrapper" match="Table/@ss:ExpandedColumnCount | Table/@ss:ExpandedRowCount"/>

  <!--
    Sort the cells by their position in the master column list;
    with any luck, this will be pretty similar to the input order, though it
    is technically implementation-defined (due to the use of distinct-values() above).
    However, even if it comes out in a different order, the data and columns should
    all match up correctly.
  -->
  <xsl:template match="Row">
    <xsl:copy>
      <xsl:apply-templates select="@*"/>
      <xsl:apply-templates select="Cell">
        <xsl:sort select="my:master-position-of(.)"/>
      </xsl:apply-templates>
    </xsl:copy>
  </xsl:template>

  <!-- For simplicity's sake, all we (need to) do is add ss:Index to every cell
       based on that cell's position in the master column list -->
  <xsl:template match="Cell">
    <xsl:copy>
      <xsl:apply-templates select="@*"/>
      <xsl:attribute name="ss:Index" select="my:master-position-of(.)"/>
      <xsl:apply-templates/>
    </xsl:copy>
  </xsl:template>

  <!-- This function gets what the column position of a cell will be in the final, merged spreadsheet -->
  <xsl:function name="my:master-position-of" as="xs:integer">
    <xsl:param name="cell" as="element(Cell)"/>
    <xsl:sequence select="index-of($master-column-keys, my:column-key(my:column-heading($cell)))"/>
  </xsl:function>

  <!-- This function gets the column heading cell for a given data cell -->
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

  <!-- This helper function gets the physical position of a Cell element within its parent -->
  <xsl:function name="my:position" as="xs:integer">
    <xsl:param name="cell" as="element(Cell)"/>
    <xsl:sequence select="1 + count($cell/preceding-sibling::Cell)"/>
  </xsl:function>

  <!-- This function gets the column key that we use internally to uniquely identify columns -->
  <xsl:function name="my:column-key" as="xs:string">
    <xsl:param name="heading" as="element(Cell)"/>
    <xsl:sequence select="concat($heading/Data,
                                 $column-key-separator,
                                 1 + count($heading/preceding-sibling::Cell[Data eq $heading/Data])
                                )"/>
  </xsl:function>

  <!-- ASSUMPTION: Your column headings do not have this string in them -->
  <xsl:variable name="column-key-separator" select="'____$!@#$____'"/>

</xsl:stylesheet>
