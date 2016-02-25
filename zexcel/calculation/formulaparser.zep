namespace ZExcel\Calculation;

class FormulaParser
{
    /* Character constants */
    const QUOTE_DOUBLE  = "\"";
    const QUOTE_SINGLE  = "\'";
    const BRACKET_CLOSE = "]";
    const BRACKET_OPEN  = "[";
    const BRACE_OPEN    = "{";
    const BRACE_CLOSE   = "}";
    const PAREN_OPEN    = "(";
    const PAREN_CLOSE   = ")";
    const SEMICOLON     = ";";
    const WHITESPACE    = " ";
    const COMMA         = ",";
    const ERROR_START   = "#";

    const OPERATORS_SN      = "+-";
    const OPERATORS_INFIX   = "+-*/^&=><";
    const OPERATORS_POSTFIX = "%";

    /**
     * Formula
     *
     * @var string
     */
    private formula;

    /**
     * Tokens
     *
     * @var \ZExcel\Calculation\FormulaToken[]
     */
    private tokens = [];

    /**
     * Create a new \ZExcel\Calculation\FormulaParser
     *
     * @param     string        pFormula    Formula to parse
     * @throws     \ZExcel\Calculation\Exception
     */
    public function __construct(string pFormula = "")
    {
        // Check parameters
        if (is_null(pFormula)) {
            throw new \ZExcel\Calculation\Exception("Invalid parameter passed: formula");
        }

        // Initialise values
        let this->formula = trim(pFormula);
        // Parse!
        this->parseToTokens();
    }

    /**
     * Get Formula
     *
     * @return string
     */
    public function getFormula()
    {
        return this->formula;
    }

    /**
     * Get Token
     *
     * @param     int        pId    Token id
     * @return    string
     * @throws  \ZExcel\Calculation\Exception
     */
    public function getToken(pId = 0)
    {
        if (isset(this->tokens[pId])) {
            return this->tokens[pId];
        } else {
            throw new \ZExcel\Calculation\Exception("Token with id pId does not exist.");
        }
    }

    /**
     * Get Token count
     *
     * @return string
     */
    public function getTokenCount() -> int
    {
        return count(this->tokens);
    }

    /**
     * Get Tokens
     *
     * @return \ZExcel\Calculation\FormulaToken[]
     */
    public function getTokens()
    {
        return this->tokens;
    }

    /**
     * Parse to tokens
     */
    private function parseToTokens()
    {
        throw new \Exception("Not implemented yet!");
    }
}
