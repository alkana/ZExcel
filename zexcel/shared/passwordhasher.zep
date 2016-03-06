namespace ZExcel\Shared;

class PasswordHasher
{
    /**
     * Create a password hash from a given string.
     *
     * This method is based on the algorithm provided by
     * Daniel Rentz of OpenOffice and the PEAR package
     * Spreadsheet_Excel_Writer by Xavier Noguer <xnoguer@rezebra.com>.
     *
     * @param     string    pPassword    Password to hash
     * @return     string                Hashed password
     */
    public static function hashPassword(pPassword = "") {
    	var password, charPos, chars, chr, value, rotated_bits;
    	
        let password    = 0x0000;
        let charPos    = 1;       // char position

        // split the plain text password in its component characters
        let chars = preg_split("//", pPassword, -1, PREG_SPLIT_NO_EMPTY);
        for chr in chars {
            let value        = ord(chr) << charPos;    // shifted ASCII value
        	let charPos      = charPos + 1;
            let rotated_bits = value >> 15;                // rotated bits beyond bit 15
            let value        = value & 0x7fff;                    // first 15 bits
            let password     = password ^ (value | rotated_bits);
        }

        let password = password ^ strlen(pPassword);
        let password = password ^ 0xCE4B;

        return (strtoupper(dechex(password)));
    }
}
