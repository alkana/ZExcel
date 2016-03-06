namespace ZExcel;

class HashTable
{
    /**
     * HashTable elements
     *
     * @var array
     */
    public _items = [];

    /**
     * HashTable key map
     *
     * @var array
     */
    public _keyMap = [];

    /**
     * Create a new \ZExcel\HashTable
     *
     * @param    \ZExcel\IComparable[] pSource    Optional source array to create HashTable from
     * @throws    \ZExcel\Exception
     */
    public function __construct(pSource = null)
    {
        if (pSource !== null) {
            // Create HashTable
            this->addFromSource(pSource);
        }
    }

    /**
     * Add HashTable items from source
     *
     * @param    \ZExcel\IComparable[] pSource    Source array to create HashTable from
     * @throws    \ZExcel\Exception
     */
    public function addFromSource(pSource = null)
    {
    	var item;
    	
        // Check if an array was passed
        if (pSource == null) {
            return null;
        } elseif (!is_array(pSource)) {
            throw new \ZExcel\Exception("Invalid array parameter passed.");
        }

        for item in pSource {
            this->add(item);
        }
    }

    /**
     * Add HashTable item
     *
     * @param    \ZExcel\IComparable pSource    Item to add
     * @throws    \ZExcel\Exception
     */
    public function add(<\ZExcel\IComparable> pSource = null)
    {
    	var hash;
    	
        let hash = pSource->getHashCode();
        
        if (!isset(this->_items[hash])) {
            let this->_items[hash] = pSource;
            let this->_keyMap[count(this->_items) - 1] = hash;
        }
    }

    /**
     * Remove HashTable item
     *
     * @param    \ZExcel\IComparable pSource    Item to remove
     * @throws    \ZExcel\Exception
     */
    public function remove(<\ZExcel\IComparable> pSource = null)
    {
    	var hash, deleteKey, key, value;
    	
        let hash = pSource->getHashCode();
        
        if (isset(this->_items[hash])) {
            unset(this->_items[hash]);

            let deleteKey = -1;
            
            for key, value in this->_keyMap {
                if (deleteKey >= 0) {
                    let this->_keyMap[key - 1] = value;
                }

                if (value == hash) {
                    let deleteKey = key;
                }
            }
            
            unset(this->_keyMap[count(this->_keyMap) - 1]);
        }
    }

    /**
     * Clear HashTable
     *
     */
    public function clear()
    {
        let this->_items = [];
        let this->_keyMap = [];
    }

    /**
     * Count
     *
     * @return int
     */
    public function count()
    {
        return count(this->_items);
    }

    /**
     * Get index for hash code
     *
     * @param    string    pHashCode
     * @return    int    Index
     */
    public function getIndexForHashCode(var pHashCode = "")
    {
        return array_search(pHashCode, this->_keyMap);
    }

    /**
     * Get by index
     *
     * @param    int    pIndex
     * @return    \ZExcel\IComparable
     *
     */
    public function getByIndex(var pIndex = 0)
    {
        if (isset(this->_keyMap[pIndex])) {
            return this->getByHashCode( this->_keyMap[pIndex] );
        }

        return null;
    }

    /**
     * Get by hashcode
     *
     * @param    string    pHashCode
     * @return    \ZExcel\IComparable
     *
     */
    public function getByHashCode(var pHashCode = "")
    {
        if (isset(this->_items[pHashCode])) {
            return this->_items[pHashCode];
        }

        return null;
    }

    /**
     * HashTable to array
     *
     * @return \ZExcel\IComparable[]
     */
    public function toArray()
    {
        return this->_items;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
    	var vars, key, value;
    	
        let vars = get_object_vars(this);
        
        for key, value in vars {
            if (is_object(value)) {
                let this->{key} = clone value;
            }
        }
    }
}
