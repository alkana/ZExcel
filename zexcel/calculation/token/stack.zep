namespace ZExcel\Calculation\Token;

class Stack
{
    /**
     *  The parser stack for formulae
     *
     *  @var mixed[]
     */
    private stack = [];

    /**
     *  Count of entries in the parser stack
     *
     *  @var integer
     */
    private count = 0;

    /**
     * Return the number of entries on the stack
     *
     * @return  integer
     */
    public function count()
    {
        return this->count;
    }

    /**
     * Push a new entry onto the stack
     *
     * @param  mixed  type
     * @param  mixed  value
     * @param  mixed  reference
     */
    public function push(var type, var value, var reference = null)
    {
        var localeFunction;
        
        let this->stack[this->count] = [
            "type": type,
            "value": value,
            "reference": reference
        ];
        
        if (type == "Function") {
            let localeFunction = \ZExcel\Calculation::_localeFunc(value);
            
            if (localeFunction != value) {
                let this->stack[this->count]["localeValue"] = localeFunction;
            }
        }
        
        let this->count = this->count + 1;
    }

    /**
     * Pop the last entry from the stack
     *
     * @return  mixed
     */
    public function pop()
    {
        if (this->count > 0) {
            let this->count = this->count - 1;
            
            return this->stack[this->count];
        }
        
        return null;
    }

    /**
     * Return an entry from the stack without removing it
     *
     * @param   integer  n  number indicating how far back in the stack we want to look
     * @return  mixed
     */
    public function last(int n = 1)
    {
        if (this->count - n < 0) {
            return null;
        }
        
        return this->stack[this->count - n];
    }

    /**
     * Clear the stack
     */
    public function clear()
    {
        let this->stack = [];
        let this->count = 0;
    }
}
