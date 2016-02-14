namespace ZExcel\RichText;

interface ITextElement
{
    /**
     * Get text
     *
     * @return string    Text
     */
    public function getText();

    /**
     * Set text
     *
     * @param     $pText string    Text
     * @return PHPExcel_RichText_ITextElement
     */
    public function setText(string pText = "");

    /**
     * Get font
     *
     * @return PHPExcel_Style_Font
     */
    public function getFont();

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode();
}
