namespace ZExcel\Shared;

class Drawing
{
    /**
     * Convert pixels to EMU
     *
     * @param     int pValue    Value in pixels
     * @return     int            Value in EMU
     */
    public static function pixelsToEMU(int pValue = 0)
    {
        return round(pValue * 9525);
    }

    /**
     * Convert EMU to pixels
     *
     * @param     int pValue    Value in EMU
     * @return     int            Value in pixels
     */
    public static function EMUToPixels(double pValue = 0)
    {
        if (pValue != 0) {
            return round(pValue / 9525);
        } else {
            return 0;
        }
    }

    /**
     * Convert pixels to column width. Exact algorithm not known.
     * By inspection of a real Excel file using Calibri 11, one finds 1000px ~ 142.85546875
     * This gives a conversion factor of 7. Also, we assume that pixels and font size are proportional.
     *
     * @param     int pValue    Value in pixels
     * @param     \ZExcel\Style\Font pDefaultFont    Default font of the workbook
     * @return     int            Value in cell dimension
     */
    public static function pixelsToCellDimension(var pValue = 0, <\ZExcel\Style\Font> pDefaultFont)
    {
        var name, size, colWidth;
        
        // Font name and size
        let name = pDefaultFont->getName();
        let size = pDefaultFont->getSize();

        if (isset(\ZExcel\Shared\Font::defaultColumnWidths[name][size])) {
            // Exact width can be determined
            let colWidth = pValue * \ZExcel\Shared\Font::defaultColumnWidths[name][size]["width"] / \ZExcel\Shared\Font::defaultColumnWidths[name][size]["px"];
        } else {
            // We don"t have data for this particular font and size, use approximation by
            // extrapolating from Calibri 11
            let colWidth = pValue * 11 * \ZExcel\Shared\Font::defaultColumnWidths["Calibri"][11]["width"] / \ZExcel\Shared\Font::defaultColumnWidths["Calibri"][11]["px"] / size;
        }

        return colWidth;
    }

    /**
     * Convert column width from (intrinsic) Excel units to pixels
     *
     * @param     float    pValue        Value in cell dimension
     * @param     \ZExcel\Style\Font pDefaultFont    Default font of the workbook
     * @return     int        Value in pixels
     */
    public static function cellDimensionToPixels(var pValue = 0, <\ZExcel\Style\Font> pDefaultFont)
    {
        var name, size, colWidth;
        
        // Font name and size
        let name = pDefaultFont->getName();
        let size = pDefaultFont->getSize();

        if (isset(\ZExcel\Shared\Font::defaultColumnWidths[name][size])) {
            // Exact width can be determined
            let colWidth = pValue * \ZExcel\Shared\Font::defaultColumnWidths[name][size]["px"] / \ZExcel\Shared\Font::defaultColumnWidths[name][size]["width"];
        } else {
            // We don"t have data for this particular font and size, use approximation by
            // extrapolating from Calibri 11
            let colWidth = pValue * size * \ZExcel\Shared\Font::defaultColumnWidths["Calibri"][11]["px"] / \ZExcel\Shared\Font::defaultColumnWidths["Calibri"][11]["width"] / 11;
        }

        // Round pixels to closest integer
        let colWidth = (int) round(colWidth);

        return colWidth;
    }

    /**
     * Convert pixels to points
     *
     * @param     int pValue    Value in pixels
     * @return     int            Value in points
     */
    public static function pixelsToPoints(double pValue = 0)
    {
        return pValue * 0.67777777;
    }

    /**
     * Convert points to pixels
     *
     * @param     int pValue    Value in points
     * @return     int            Value in pixels
     */
    public static function pointsToPixels(double pValue = 0)
    {
        if (pValue != 0) {
            return (int) ceil(pValue * 1.333333333);
        } else {
            return 0;
        }
    }

    /**
     * Convert degrees to angle
     *
     * @param     int pValue    Degrees
     * @return     int            Angle
     */
    public static function degreesToAngle(double pValue = 0)
    {
        return (int)round(pValue * 60000);
    }

    /**
     * Convert angle to degrees
     *
     * @param     int pValue    Angle
     * @return     int            Degrees
     */
    public static function angleToDegrees(double pValue = 0)
    {
        if (pValue != 0) {
            return round(pValue / 60000);
        } else {
            return 0;
        }
    }

    /**
     * Create a new image from file. By alexander at alexauto dot nl
     *
     * @link http://www.php.net/manual/en/function.imagecreatefromwbmp.php#86214
     * @param string filename Path to Windows DIB (BMP) image
     * @return resource
     */
    public static function imagecreatefrombmp(var p_sFile)
    {
        var file, read, temp, hex, header, header_parts, header_size,
            width, height, b, i, x, y, g, r, image, color,
            body, body_size, usePadding, i_pos;
        
        //    Load the image into a string
        let file = fopen(p_sFile, "rb");
        let read = fread(file, 10);
        
        while (!feof(file) && (read != "")) {
            let read = read . fread(file, 1024);
        }

        let temp = unpack("H*", read);
        let hex = temp[1];
        let header = substr(hex, 0, 108);

        //    Process the header
        //    Structure: http://www.fastgraph.com/help/bmp_header_format.html
        if (substr(header, 0, 4) == "424d") {
            //    Cut it in parts of 2 bytes
            let header_parts = str_split(header, 2);

            //    Get the width        4 bytes
            let width = hexdec(header_parts[19].header_parts[18]);

            //    Get the height        4 bytes
            let height = hexdec(header_parts[23].header_parts[22]);
        }

        //    Define starting X and Y
        let x = 0;
        let y = 1;

        //    Create newimage
        let image = imagecreatetruecolor(width, height);

        //    Grab the body from the image
        let body = substr(hex, 108);

        //    Calculate if padding at the end-line is needed
        //    Divided by two to keep overview.
        //    1 byte = 2 HEX-chars
        let body_size = (strlen(body) / 2);
        let header_size = (width * height);

        //    Use end-line padding? Only when needed
        let usePadding = (body_size > (header_size * 3) + 4);

        //    Using a for-loop with index-calculation instaid of str_split to avoid large memory consumption
        //    Calculate the next DWORD-position in the body
        let i = 0;
        
        while i < body_size {
            //    Calculate line-ending and padding
            if (x >= width) {
                // If padding needed, ignore image-padding
                // Shift i to the ending of the current 32-bit-block
                if (usePadding) {
                    let i = i + (width % 4);
                }

                //    Reset horizontal position
                let x = 0;

                //    Raise the height-position (bottom-up)
                let y = y + 1;

                //    Reached the image-height? Break the for-loop
                if (y > height) {
                    break;
                }
            }

            // Calculation of the RGB-pixel (defined as BGR in image-data)
            // Define i_pos as absolute position in the body
            let i_pos    = i * 2;
            let r        = hexdec(body[i_pos + 4] . body[i_pos + 5]);
            let g        = hexdec(body[i_pos + 2] . body[i_pos + 3]);
            let b        = hexdec(body[i_pos] . body[i_pos + 1]);

            // Calculate and draw the pixel
            let color = imagecolorallocate(image, r, g, b);
            imagesetpixel(image, x, height - y, color);

            // Raise the horizontal position
            let x = x + 1;
            let i = i + 3;
        }

        //    Return image-object
        return image;
    }
}
