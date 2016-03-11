namespace ZExcel;

class Chart
{
    /**
     * Chart Name
     *
     * @var string
     */
    private name = "";

    /**
     * Worksheet
     *
     * @var \ZExcel\Worksheet
     */
    private worksheet;

    /**
     * Chart Title
     *
     * @var \ZExcel\Chart\Title
     */
    private title;

    /**
     * Chart Legend
     *
     * @var \ZExcel\Chart\Legend
     */
    private legend;

    /**
     * X-Axis Label
     *
     * @var \ZExcel\Chart\Title
     */
    private xAxisLabel;

    /**
     * Y-Axis Label
     *
     * @var \ZExcel\Chart\Title
     */
    private yAxisLabel;

    /**
     * Chart Plot Area
     *
     * @var \ZExcel\Chart\PlotArea
     */
    private plotArea;

    /**
     * Plot Visible Only
     *
     * @var boolean
     */
    private plotVisibleOnly = true;

    /**
     * Display Blanks as
     *
     * @var string
     */
    private displayBlanksAs = "0";

    /**
     * Chart Asix Y as
     *
     * @var \ZExcel\Chart\Axis
     */
    private yAxis;

    /**
     * Chart Asix X as
     *
     * @var \ZExcel\Chart\Axis
     */
    private xAxis;

    /**
     * Chart Major Gridlines as
     *
     * @var \ZExcel\Chart\GridLines
     */
    private majorGridlines;

    /**
     * Chart Minor Gridlines as
     *
     * @var \ZExcel\Chart\GridLines
     */
    private minorGridlines;

    /**
     * Top-Left Cell Position
     *
     * @var string
     */
    private topLeftCellRef = "A1";


    /**
     * Top-Left X-Offset
     *
     * @var integer
     */
    private topLeftXOffset = 0;


    /**
     * Top-Left Y-Offset
     *
     * @var integer
     */
    private topLeftYOffset = 0;


    /**
     * Bottom-Right Cell Position
     *
     * @var string
     */
    private bottomRightCellRef = "A1";


    /**
     * Bottom-Right X-Offset
     *
     * @var integer
     */
    private bottomRightXOffset = 10;


    /**
     * Bottom-Right Y-Offset
     *
     * @var integer
     */
    private bottomRightYOffset = 10;


    /**
     * Create a new \ZExcel\Chart
     */
    public function __construct(name, <\ZExcel\Chart\Title> title = null, <\ZExcel\Chart\Legend> legend = null, <\ZExcel\Chart\PlotArea> plotArea = null, plotVisibleOnly = true, displayBlanksAs = "0", <\ZExcel\Chart\Title> xAxisLabel = null, <\ZExcel\Chart\Title> yAxisLabel = null, <\ZExcel\Chart\Axis> xAxis = null, <\ZExcel\Chart\Axis> yAxis = null, <\ZExcel\Chart\GridLines> majorGridlines = null, <\ZExcel\Chart\GridLines> minorGridlines = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get Name
     *
     * @return string
     */
    public function getName()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get Worksheet
     *
     * @return \ZExcel\Worksheet
     */
    public function getWorksheet()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set Worksheet
     *
     * @param    \ZExcel\Worksheet    pValue
     * @throws    \ZExcel\Chart\Exception
     * @return \ZExcel\Chart
     */
    public function setWorksheet(<\ZExcel\Worksheet> pValue = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get Title
     *
     * @return \ZExcel\Chart\Title
     */
    public function getTitle()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set Title
     *
     * @param    \ZExcel\Chart\Title title
     * @return    \ZExcel\Chart
     */
    public function setTitle(<\ZExcel\Chart\Title> title)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get Legend
     *
     * @return \ZExcel\Chart\Legend
     */
    public function getLegend()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set Legend
     *
     * @param    \ZExcel\Chart\Legend legend
     * @return    \ZExcel\Chart
     */
    public function setLegend(<\ZExcel\Chart\Legend> legend)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get X-Axis Label
     *
     * @return \ZExcel\Chart\Title
     */
    public function getXAxisLabel()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set X-Axis Label
     *
     * @param    \ZExcel\Chart\Title label
     * @return    \ZExcel\Chart
     */
    public function setXAxisLabel(<\ZExcel\Chart\Title> label)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get Y-Axis Label
     *
     * @return \ZExcel\Chart\Title
     */
    public function getYAxisLabel()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set Y-Axis Label
     *
     * @param    \ZExcel\Chart\Title label
     * @return    \ZExcel\Chart
     */
    public function setYAxisLabel(<\ZExcel\Chart\Title> label)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get Plot Area
     *
     * @return \ZExcel\Chart\PlotArea
     */
    public function getPlotArea()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get Plot Visible Only
     *
     * @return boolean
     */
    public function getPlotVisibleOnly()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set Plot Visible Only
     *
     * @param boolean plotVisibleOnly
     * @return \ZExcel\Chart
     */
    public function setPlotVisibleOnly(plotVisibleOnly = true)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get Display Blanks as
     *
     * @return string
     */
    public function getDisplayBlanksAs()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set Display Blanks as
     *
     * @param string displayBlanksAs
     * @return \ZExcel\Chart
     */
    public function setDisplayBlanksAs(displayBlanksAs = "0")
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     * Get yAxis
     *
     * @return \ZExcel\Chart\Axis
     */
    public function getChartAxisY()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get xAxis
     *
     * @return \ZExcel\Chart\Axis
     */
    public function getChartAxisX()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get Major Gridlines
     *
     * @return \ZExcel\Chart\GridLines
     */
    public function getMajorGridlines()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get Minor Gridlines
     *
     * @return \ZExcel\Chart\GridLines
     */
    public function getMinorGridlines()
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     * Set the Top Left position for the chart
     *
     * @param    string    cell
     * @param    integer    xOffset
     * @param    integer    yOffset
     * @return \ZExcel\Chart
     */
    public function setTopLeftPosition(cell, xOffset = null, yOffset = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get the top left position of the chart
     *
     * @return array    an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
     */
    public function getTopLeftPosition()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get the cell address where the top left of the chart is fixed
     *
     * @return string
     */
    public function getTopLeftCell()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set the Top Left cell position for the chart
     *
     * @param    string    cell
     * @return \ZExcel\Chart
     */
    public function setTopLeftCell(cell)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set the offset position within the Top Left cell for the chart
     *
     * @param    integer    xOffset
     * @param    integer    yOffset
     * @return \ZExcel\Chart
     */
    public function setTopLeftOffset(xOffset = null, yOffset = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get the offset position within the Top Left cell for the chart
     *
     * @return integer[]
     */
    public function getTopLeftOffset()
    {
        throw new \Exception("Not implemented yet!");
    }

    public function setTopLeftXOffset(xOffset)
    {
        throw new \Exception("Not implemented yet!");
    }

    public function getTopLeftXOffset()
    {
        throw new \Exception("Not implemented yet!");
    }

    public function setTopLeftYOffset(yOffset)
    {
        throw new \Exception("Not implemented yet!");
    }

    public function getTopLeftYOffset()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set the Bottom Right position of the chart
     *
     * @param    string    cell
     * @param    integer    xOffset
     * @param    integer    yOffset
     * @return \ZExcel\Chart
     */
    public function setBottomRightPosition(cell, xOffset = null, yOffset = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get the bottom right position of the chart
     *
     * @return array    an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
     */
    public function getBottomRightPosition()
    {
        throw new \Exception("Not implemented yet!");
    }

    public function setBottomRightCell(cell)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get the cell address where the bottom right of the chart is fixed
     *
     * @return string
     */
    public function getBottomRightCell()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set the offset position within the Bottom Right cell for the chart
     *
     * @param    integer    xOffset
     * @param    integer    yOffset
     * @return \ZExcel\Chart
     */
    public function setBottomRightOffset(xOffset = null, yOffset = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get the offset position within the Bottom Right cell for the chart
     *
     * @return integer[]
     */
    public function getBottomRightOffset()
    {
        throw new \Exception("Not implemented yet!");
    }

    public function setBottomRightXOffset(xOffset)
    {
        throw new \Exception("Not implemented yet!");
    }

    public function getBottomRightXOffset()
    {
        throw new \Exception("Not implemented yet!");
    }

    public function setBottomRightYOffset(yOffset)
    {
        throw new \Exception("Not implemented yet!");
    }

    public function getBottomRightYOffset()
    {
        throw new \Exception("Not implemented yet!");
    }


    public function refresh()
    {
        throw new \Exception("Not implemented yet!");
    }

    public function render(outputDestination = null)
    {
        throw new \Exception("Not implemented yet!");
    }
}
