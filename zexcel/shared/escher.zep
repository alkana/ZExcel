namespace ZExcel\Shared;

class Escher
{
    /**
     * Drawing Group Container
     *
     * @var \ZExcel\Shared\Escher\DggContainer
     */
    private dggContainer;

    /**
     * Drawing Container
     *
     * @var \ZExcel\Shared\Escher\DgContainer
     */
    private dgContainer;

    /**
     * Get Drawing Group Container
     *
     * @return \ZExcel\Shared\Escher\DgContainer
     */
    public function getDggContainer()
    {
        return this->dggContainer;
    }

    /**
     * Set Drawing Group Container
     *
     * @param \ZExcel\Shared\Escher\DggContainer dggContainer
     */
    public function setDggContainer(var dggContainer)
    {
        let this->dggContainer = dggContainer;
        
        return this->dggContainer;
    }

    /**
     * Get Drawing Container
     *
     * @return \ZExcel\Shared\Escher\DgContainer
     */
    public function getDgContainer()
    {
        return this->dgContainer;
    }

    /**
     * Set Drawing Container
     *
     * @param \ZExcel\Shared\Escher\DgContainer dgContainer
     */
    public function setDgContainer(var dgContainer)
    {
        let this->dgContainer = dgContainer;
        
        return this->dgContainer;
    }
}
