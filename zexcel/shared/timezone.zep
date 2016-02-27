namespace ZExcel\Shared;

class TimeZone
{
    /**
     * Default Timezone used for date/time conversions
     *
     * @private
     * @var    string
     */
    protected static _timezone    = "UTC";

    /**
     * Validate a Timezone name
     *
     * @param  string  timezone Time zone (e.g. "Europe/London")
     * @return boolean Success or failure
     */
    public static function _validateTimeZone(string timezone) {
        if (in_array(timezone, \DateTimeZone::listIdentifiers())) {
            return true;
        }
        return false;
    }

    /**
     * Set the Default Timezone used for date/time conversions
     *
     * @param     string        timezone            Time zone (e.g. "Europe/London")
     * @return     boolean                        Success or failure
     */
    public static function setTimeZone(timezone) -> boolean
    {
        if (self::_validateTimezone(timezone)) {
            let self::_timezone = timezone;
            return true;
        }
        return false;
    }    //    function setTimezone()


    /**
     * Return the Default Timezone used for date/time conversions
     *
     * @return     string        Timezone (e.g. "Europe/London")
     */
    public static function getTimeZone() -> string
    {
        return self::_timezone;
    }    //    function getTimezone()


    /**
     *    Return the Timezone transition for the specified timezone and timestamp
     *
     *    @param  DateTimeZone objTimezone The timezone for finding the transitions
     *    @param  int          timestamp   PHP date/time value for finding the current transition
     *    @return array        The current transition details
     */
    private static function _getTimezoneTransitions(<\DateTimeZone> objTimezone, int timestamp) -> array
    {
        var allTransitions = objTimezone->getTransitions(), transitions = [], key, transition;
        
        for key, transition in allTransitions {
            if (transition["ts"] > timestamp) {
                let transitions[] = (key > 0) ? allTransitions[key - 1] : transition;
                break;
            }
            if (empty(transitions)) {
                let transitions[] = end(allTransitions);
            }
        }

        return transitions;
    }

    /**
     *    Return the Timezone offset used for date/time conversions to/from UST
     *    This requires both the timezone and the calculated date/time to allow for local DST
     *
     *    @param        string                 timezone        The timezone for finding the adjustment to UST
     *    @param        integer                 timestamp        PHP date/time value
     *    @return         integer                Number of seconds for timezone adjustment
     *    @throws        PHPExcel_Exception
     */
    public static function getTimeZoneAdjustment(var timezone, int timestamp) {
        var objTimezone, transitions;
        
        if (timezone !== null) {
            if (!self::_validateTimezone(timezone)) {
                throw new \ZExcel\Exception("Invalid timezone " . timezone);
            }
        } else {
            let timezone = self::_timezone;
        }

        if (timezone == "UST") {
            return 0;
        }

        let objTimezone = new \DateTimeZone(timezone);
        let transitions = objTimezone->getTransitions(timestamp,timestamp);
        
        if (count(transitions) > 0) {
            return transitions[0]["offset"];
        }
        
        return 0;
    }
}
