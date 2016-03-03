namespace ZExcel\Cell;

class Hyperlink
{
    /**
     * URL to link the cell to
     *
     * @var string
     */
    private url;

    /**
     * Tooltip to display on the hyperlink
     *
     * @var string
     */
    private tooltip;

    /**
     * Create a new PHPExcel_Cell_Hyperlink
     *
     * @param  string  $pUrl      Url to link the cell to
     * @param  string  $pTooltip  Tooltip to display on the hyperlink
     */
    public function __construct(string pUrl = "", string pTooltip = "")
    {
        // Initialise member variables
        let this->url     = pUrl;
        let this->tooltip = pTooltip;
    }

    /**
     * Get URL
     *
     * @return string
     */
    public function getUrl() -> string
    {
        return this->url;
    }

    /**
     * Set URL
     *
     * @param  string    $value
     * @return PHPExcel_Cell_Hyperlink
     */
    public function setUrl(string value = "") -> <\ZExcel\Cell\Hyperlink>
    {
        let this->url = value;
        return this;
    }

    /**
     * Get tooltip
     *
     * @return string
     */
    public function getTooltip() -> string
    {
        return this->tooltip;
    }

    /**
     * Set tooltip
     *
     * @param  string    $value
     * @return PHPExcel_Cell_Hyperlink
     */
    public function setTooltip(string value = "") -> <\ZExcel\Cell\Hyperlink>
    {
        let this->tooltip = value;
        return this;
    }

    /**
     * Is this hyperlink internal? (to another worksheet)
     *
     * @return boolean
     */
    public function isInternal() -> boolean
    {
        return strpos(this->url, "sheet://") !== false;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode() -> string
    {
        return md5(this->url . this->tooltip . get_class(this));
    }
}
