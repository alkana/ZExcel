namespace ZExcel\Style;

abstract class Supervisor
{
	/**
     * Supervisor?
     *
     * @var boolean
     */
    protected isSupervisor = false;

    /**
     * Parent. Only used for supervisor
     *
     * @var PHPExcel_Style
     */
    protected parent;

    /**
     * Create a new PHPExcel_Style_Alignment
     *
     * @param    boolean    isSupervisor    Flag indicating if this is a supervisor or not
     *                                    Leave this value at default unless you understand exactly what
     *                                        its ramifications are
     */
    public function __construct(boolean isSupervisor = false)
    {
        // Supervisor?
        let this->isSupervisor = isSupervisor;
    }

    /**
     * Bind parent. Only used for supervisor
     *
     * @param PHPExcel parent
     * @return PHPExcel_Style_Supervisor
     */
    public function bindParent(parent, parentPropertyName = null)
    {
        let this->parent = parent;
        return this;
    }

    /**
     * Is this a supervisor or a cell style component?
     *
     * @return boolean
     */
    public function getIsSupervisor() -> boolean
    {
        return this->isSupervisor;
    }

    /**
     * Get the currently active sheet. Only used for supervisor
     *
     * @return PHPExcel_Worksheet
     */
    public function getActiveSheet()
    {
        return this->parent->getActiveSheet();
    }

    /**
     * Get the currently active cell coordinate in currently active sheet.
     * Only used for supervisor
     *
     * @return string E.g. 'A1'
     */
    public function getSelectedCells()
    {
        return this->getActiveSheet()->getSelectedCells();
    }

    /**
     * Get the currently active cell coordinate in currently active sheet.
     * Only used for supervisor
     *
     * @return string E.g. 'A1'
     */
    public function getActiveCell()
    {
        return this->getActiveSheet()->getActiveCell();
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
    	var vars, key, value;
    	
        let vars = get_object_vars(this);
        
        for key, value in vars {
            if ((is_object(value)) && (key != "parent")) {
                let this->key = clone value;
            } else {
                let this->key = value;
            }
        }
    }
}
