namespace ZExcel\Calculation;

class Engineering
{
    const EULER = 2.71828182845904523536;

    /**
    /* Private method to calculate the erfc value
     */
    private static oneSqrtPi = 0.564189583547756287;
    
    /**
    /* Private method to calculate the erf value
     */
    private static twoSqrtPi = 1.128379167095512574;
    
    /**
     * Details of the Units of measure that can be used in CONVERTUOM()
     *
     * @var mixed[]
     */
    private static conversionUnits = [
        "g"     : ["Group" : "Mass",        "Unit Name" : "Gram",                     "AllowPrefix" : true],
        "sg"    : ["Group" : "Mass",        "Unit Name" : "Slug",                     "AllowPrefix" : false],
        "lbm"   : ["Group" : "Mass",        "Unit Name" : "Pound mass (avoirdupois)", "AllowPrefix" : false],
        "u"     : ["Group" : "Mass",        "Unit Name" : "U (atomic mass unit)",     "AllowPrefix" : true],
        "ozm"   : ["Group" : "Mass",        "Unit Name" : "Ounce mass (avoirdupois)", "AllowPrefix" : false],
        "m"     : ["Group" : "Distance",    "Unit Name" : "Meter",                    "AllowPrefix" : true],
        "mi"    : ["Group" : "Distance",    "Unit Name" : "Statute mile",             "AllowPrefix" : false],
        "Nmi"   : ["Group" : "Distance",    "Unit Name" : "Nautical mile",            "AllowPrefix" : false],
        "in"    : ["Group" : "Distance",    "Unit Name" : "Inch",                     "AllowPrefix" : false],
        "ft"    : ["Group" : "Distance",    "Unit Name" : "Foot",                     "AllowPrefix" : false],
        "yd"    : ["Group" : "Distance",    "Unit Name" : "Yard",                     "AllowPrefix" : false],
        "ang"   : ["Group" : "Distance",    "Unit Name" : "Angstrom",                 "AllowPrefix" : true],
        "Pica"  : ["Group" : "Distance",    "Unit Name" : "Pica (1/72 in)",           "AllowPrefix" : false],
        "yr"    : ["Group" : "Time",        "Unit Name" : "Year",                     "AllowPrefix" : false],
        "day"   : ["Group" : "Time",        "Unit Name" : "Day",                      "AllowPrefix" : false],
        "hr"    : ["Group" : "Time",        "Unit Name" : "Hour",                     "AllowPrefix" : false],
        "mn"    : ["Group" : "Time",        "Unit Name" : "Minute",                   "AllowPrefix" : false],
        "sec"   : ["Group" : "Time",        "Unit Name" : "Second",                   "AllowPrefix" : true],
        "Pa"    : ["Group" : "Pressure",    "Unit Name" : "Pascal",                   "AllowPrefix" : true],
        "p"     : ["Group" : "Pressure",    "Unit Name" : "Pascal",                   "AllowPrefix" : true],
        "atm"   : ["Group" : "Pressure",    "Unit Name" : "Atmosphere",               "AllowPrefix" : true],
        "at"    : ["Group" : "Pressure",    "Unit Name" : "Atmosphere",               "AllowPrefix" : true],
        "mmHg"  : ["Group" : "Pressure",    "Unit Name" : "mm of Mercury",            "AllowPrefix" : true],
        "N"     : ["Group" : "Force",       "Unit Name" : "Newton",                   "AllowPrefix" : true],
        "dyn"   : ["Group" : "Force",       "Unit Name" : "Dyne",                     "AllowPrefix" : true],
        "dy"    : ["Group" : "Force",       "Unit Name" : "Dyne",                     "AllowPrefix" : true],
        "lbf"   : ["Group" : "Force",       "Unit Name" : "Pound force",              "AllowPrefix" : false],
        "J"     : ["Group" : "Energy",      "Unit Name" : "Joule",                    "AllowPrefix" : true],
        "e"     : ["Group" : "Energy",      "Unit Name" : "Erg",                      "AllowPrefix" : true],
        "c"     : ["Group" : "Energy",      "Unit Name" : "Thermodynamic calorie",    "AllowPrefix" : true],
        "cal"   : ["Group" : "Energy",      "Unit Name" : "IT calorie",               "AllowPrefix" : true],
        "eV"    : ["Group" : "Energy",      "Unit Name" : "Electron volt",            "AllowPrefix" : true],
        "ev"    : ["Group" : "Energy",      "Unit Name" : "Electron volt",            "AllowPrefix" : true],
        "HPh"   : ["Group" : "Energy",      "Unit Name" : "Horsepower-hour",          "AllowPrefix" : false],
        "hh"    : ["Group" : "Energy",      "Unit Name" : "Horsepower-hour",          "AllowPrefix" : false],
        "Wh"    : ["Group" : "Energy",      "Unit Name" : "Watt-hour",                "AllowPrefix" : true],
        "wh"    : ["Group" : "Energy",      "Unit Name" : "Watt-hour",                "AllowPrefix" : true],
        "flb"   : ["Group" : "Energy",      "Unit Name" : "Foot-pound",               "AllowPrefix" : false],
        "BTU"   : ["Group" : "Energy",      "Unit Name" : "BTU",                      "AllowPrefix" : false],
        "btu"   : ["Group" : "Energy",      "Unit Name" : "BTU",                      "AllowPrefix" : false],
        "HP"    : ["Group" : "Power",       "Unit Name" : "Horsepower",               "AllowPrefix" : false],
        "h"     : ["Group" : "Power",       "Unit Name" : "Horsepower",               "AllowPrefix" : false],
        "W"     : ["Group" : "Power",       "Unit Name" : "Watt",                     "AllowPrefix" : true],
        "w"     : ["Group" : "Power",       "Unit Name" : "Watt",                     "AllowPrefix" : true],
        "T"     : ["Group" : "Magnetism",   "Unit Name" : "Tesla",                    "AllowPrefix" : true],
        "ga"    : ["Group" : "Magnetism",   "Unit Name" : "Gauss",                    "AllowPrefix" : true],
        "C"     : ["Group" : "Temperature", "Unit Name" : "Celsius",                  "AllowPrefix" : false],
        "cel"   : ["Group" : "Temperature", "Unit Name" : "Celsius",                  "AllowPrefix" : false],
        "F"     : ["Group" : "Temperature", "Unit Name" : "Fahrenheit",               "AllowPrefix" : false],
        "fah"   : ["Group" : "Temperature", "Unit Name" : "Fahrenheit",               "AllowPrefix" : false],
        "K"     : ["Group" : "Temperature", "Unit Name" : "Kelvin",                   "AllowPrefix" : false],
        "kel"   : ["Group" : "Temperature", "Unit Name" : "Kelvin",                   "AllowPrefix" : false],
        "tsp"   : ["Group" : "Liquid",      "Unit Name" : "Teaspoon",                 "AllowPrefix" : false],
        "tbs"   : ["Group" : "Liquid",      "Unit Name" : "Tablespoon",               "AllowPrefix" : false],
        "oz"    : ["Group" : "Liquid",      "Unit Name" : "Fluid Ounce",              "AllowPrefix" : false],
        "cup"   : ["Group" : "Liquid",      "Unit Name" : "Cup",                      "AllowPrefix" : false],
        "pt"    : ["Group" : "Liquid",      "Unit Name" : "U.S. Pint",                "AllowPrefix" : false],
        "us_pt" : ["Group" : "Liquid",      "Unit Name" : "U.S. Pint",                "AllowPrefix" : false],
        "uk_pt" : ["Group" : "Liquid",      "Unit Name" : "U.K. Pint",                "AllowPrefix" : false],
        "qt"    : ["Group" : "Liquid",      "Unit Name" : "Quart",                    "AllowPrefix" : false],
        "gal"   : ["Group" : "Liquid",      "Unit Name" : "Gallon",                   "AllowPrefix" : false],
        "l"     : ["Group" : "Liquid",      "Unit Name" : "Litre",                    "AllowPrefix" : true],
        "lt"    : ["Group" : "Liquid",      "Unit Name" : "Litre",                    "AllowPrefix" : true]
    ];

    /**
     * Details of the Multiplier prefixes that can be used with Units of Measure in CONVERTUOM()
     *
     * @var mixed[]
     */
    private static conversionMultipliers = [
        "Y" : ["multiplier" : 1000000000000000000000000,  "name" : "yotta"],
        "Z" : ["multiplier" : 1000000000000000000000,  "name" : "zetta"],
        "E" : ["multiplier" : 1000000000000000000,  "name" : "exa"],
        "P" : ["multiplier" : 1000000000000000,  "name" : "peta"],
        "T" : ["multiplier" : 1000000000000,  "name" : "tera"],
        "G" : ["multiplier" : 1000000000,   "name" : "giga"],
        "M" : ["multiplier" : 1000000,   "name" : "mega"],
        "k" : ["multiplier" : 1000,   "name" : "kilo"],
        "h" : ["multiplier" : 100,   "name" : "hecto"],
        "e" : ["multiplier" : 10,   "name" : "deka"],
        "d" : ["multiplier" : 0.1,  "name" : "deci"],
        "c" : ["multiplier" : 0.01,  "name" : "centi"],
        "m" : ["multiplier" : 0.001,  "name" : "milli"],
        "u" : ["multiplier" : 0.000001,  "name" : "micro"],
        "n" : ["multiplier" : 0.000000001,  "name" : "nano"],
        "p" : ["multiplier" : 0.000000000001, "name" : "pico"],
        "f" : ["multiplier" : 0.000000000000001, "name" : "femto"],
        "a" : ["multiplier" : 0.000000000000000001, "name" : "atto"],
        "z" : ["multiplier" : 0.000000000000000000001, "name" : "zepto"],
        "y" : ["multiplier" : 0.000000000000000000000001, "name" : "yocto"]
    ];

    /**
     * Details of the Units of measure conversion factors, organised by group
     *
     * @var mixed[]
     */
    public static unitConversions = [
        "Mass" : [
            "g" : [
                "g"   : 1,
                "sg"  : 0.0000685220500053478,
                "lbm" : 0.0022046229146913,
                "u"   : 602217000000000000000000,
                "ozm" : 0.035273971800363
            ],
            "sg" : [
                "g"   : 14593.8424189287,
                "sg"  : 1,
                "lbm" : 32.1739194101647,
                "u"   : 8788660000000000000000000000,
                "ozm" : 514.782785944229
            ],
            "lbm" : [
                "g"   : 453.59230974881148,
                "sg"  : 0.0310810749306493,
                "lbm" : 1,
                "u"   : 273161000000000000000000000,
                "ozm" : 16.000002342941
            ],
            "u" : [
                "g"   : 0.00000000000000000000000166053100460465,
                "sg"  : 0.00000000000000000000000000011378298853295,
                "lbm" : 0.00000000000000000000000000366084470330684,
                "u"   : 1,
                "ozm" : 0.0000000000000000000000000585735238300524
            ],
            "ozm" : [
                "g"   : 28.3495152079732,
                "sg"  : 0.00194256689870811,
                "lbm" : 0.0624999908478882,
                "u"   : 17072560000000000000000000,
                "ozm" : 1
            ]
        ],
        "Distance" : [
            "m" : [
                "m"    : 1,
                "mi"   : 0.000621371192237334,
                "Nmi"  : 0.000539956803455724,
                "in"   : 9.3700787401575,
                "ft"   : 3.28083989501312,
                "yd"   : 1.09361329797891,
                "ang"  : 10000000000,
                "Pica" : 2834.64566929116
            ],
            "mi" : [
                "m"    : 1609.344,
                "mi"   : 1,
                "Nmi"  : 0.868976241900648,
                "in"   : 63360,
                "ft"   : 5280,
                "yd"   : 1760,
                "ang"  : 16093440000000,
                "Pica" : 4561919.99999971
            ],
            "Nmi" : [
                "m"    : 1852,
                "mi"   : 1.15077944802354,
                "Nmi"  : 1,
                "in"   : 72913.3858267717,
                "ft"   : 6076.1154855643,
                "yd"   : 2025.37182785694,
                "ang"  : 18520000000000,
                "Pica" : 5249763.77952723
            ],
            "in" : [
                "m"    : 0.0254,
                "mi"   : 0.0000157828282828283,
                "Nmi"  : 0.0000137149028077754,
                "in"   : 1,
                "ft"   : 0.0833333333333333,
                "yd"   : 0.0277777777686643,
                "ang"  : 254000000,
                "Pica" : 71.9999999999955
            ],
            "ft" : [
                "m"    : 0.3048,
                "mi"   : 0.000189393939393939,
                "Nmi"  : 0.000164578833693305,
                "in"   : 10.2,
                "ft"   : 1,
                "yd"   : 0.333333333223972,
                "ang"  : 3048000000,
                "Pica" : 863.999999999946
            ],
            "yd" : [
                "m"    : 0.91440000030,
                "mi"   : 0.000568181818368230,
                "Nmi"  : 0.000493736501241901,
                "in"   : 36.000000011811,
                "ft"   : 3,
                "yd"   : 1,
                "ang"  : 9144000003,
                "Pica" : 2592.00000085023
            ],
            "ang" : [
                "m"    : 0.0000000001,
                "mi"   : 0.0000000000000621371192237334,
                "Nmi"  : 0.0000000000000539956803455724,
                "in"   : 0.00000000393700787401575,
                "ft"   : 0.000000000328083989501312,
                "yd"   : 0.000000000109361329797891,
                "ang"  : 1,
                "Pica" : 0.000000283464566929116
            ],
            "Pica" : [
                "m"    : 0.0003527777777778,
                "mi"   : 0.000000219205948372629,
                "Nmi"  : 0.000000190484761219114,
                "in"   : 0.0138888888888898,
                "ft"   : 0.00115740740740748,
                "yd"   : 0.000385802469009251,
                "ang"  : 3527777.777778,
                "Pica" : 1
            ]
        ],
        "Time" : [
            "yr" : [
                "yr"  : 1,
                "day" : 365.25,
                "hr"  : 8766,
                "mn"  : 525960,
                "sec" : 31557600
            ],
            "day" : [
                "yr"  : 0.00273785078713210,
                "day" : 1,
                "hr"  : 24,
                "mn"  : 1440,
                "sec" : 86400
            ],
            "hr" : [
                "yr"  : 0.000114077116130504,
                "day" : 0.0416666666666667,
                "hr"  : 1,
                "mn"  : 60,
                "sec" : 3600
            ],
            "mn" : [
                "yr"  : 0.00000190128526884174,
                "day" : 0.000694444444444444,
                "hr"  : 0.0166666666666667,
                "mn"  : 1,
                "sec" : 60
            ],
            "sec" : [
                "yr"  : 0.000000000000000316880878140289,
                "day" : 0.000000000115740740740741,
                "hr"  : 0.0000000277777777777778,
                "mn"  : 0.000166666666666667,
                "sec" : 1
            ]
        ],
        "Pressure" : [
            "Pa" : [
                "Pa"   : 1,
                "p"    : 1,
                "atm"  : 0.00000986923299998193,
                "at"   : 0.00000986923299998193,
                "mmHg" : 0.00750061707998627
            ],
            "p" : [
                "Pa"   : 1,
                "p"    : 1,
                "atm"  : 0.00000986923299998193,
                "at"   : 0.00000986923299998193,
                "mmHg" : 0.00750061707998627
            ],
            "atm" : [
                "Pa"   : 101324.996583,
                "p"    : 101324.996583,
                "atm"  : 1,
                "at"   : 1,
                "mmHg" : 760
            ],
            "at" : [
                "Pa"   : 101324.996583,
                "p"    : 101324.996583,
                "atm"  : 1,
                "at"   : 1,
                "mmHg" : 760
            ],
            "mmHg" : [
                "Pa"   : 133.322363925,
                "p"    : 133.322363925,
                "atm"  : 0.00131578947368421,
                "at"   : 0.00131578947368421,
                "mmHg" : 1
            ]
        ],
        "Force" : [
            "N" : [
                "N"   : 1,
                "dyn" : 100000,
                "dy"  : 100000,
                "lbf" : 0.224808923655339
            ],
            "dyn" : [
                "N"   : 100000,
                "dyn" : 1,
                "dy"  : 1,
                "lbf" : 0.00000224808923655339
            ],
            "dy" : [
                "N"   : 100000,
                "dyn" : 1,
                "dy"  : 1,
                "lbf" : 0.00000224808923655339
            ],
            "lbf" : [
                "N"   : 4.448222,
                "dyn" : 444822.2,
                "dy"  : 444822.2,
                "lbf" : 1
            ]
        ],
        "Energy" : [
            "J" : [
                "J"   : 1,
                "e"   : 9999995.19343231,
                "c"   : 0.239006249473467,
                "cal" : 0.238846190642017,
                "eV"  : 6241457000000000000,
                "ev"  : 6241457000000000000,
                "HPh" : 0.000000372506430801,
                "hh"  : 0.000000372506430801,
                "Wh"  : 0.000277777916238711,
                "wh"  : 0.000277777916238711,
                "flb" : 23.7304222192651,
                "BTU" : 0.000947815067349015,
                "btu" : 0.000947815067349015
            ],
            "e" : [
                "J"   : 0.0000001000000480657,
                "e"   : 1,
                "c"   : 0.0000000239006364353494,
                "cal" : 0.0000000238846305445111,
                "eV"  : 624146000000,
                "ev"  : 624146000000,
                "HPh" : 0.00000000000000372506609848824,
                "hh"  : 0.00000000000000372506609848824,
                "Wh"  : 0.0000000000277778049754611,
                "wh"  : 0.0000000000277778049754611,
                "flb" : 0.00000237304336254586,
                "BTU" : 0.0000000000947815522922962,
                "btu" : 0.0000000000947815522922962
            ],
            "c" : [
                "J"   : 4.18399101363672,
                "e"   : 41839890.0257312,
                "c"   : 1,
                "cal" : 0.999330315287563,
                "eV"  : 26114200000000000000,
                "ev"  : 26114200000000000000,
                "HPh" : 0.00000155856355899327,
                "hh"  : 0.00000155856355899327,
                "Wh"  : 0.0011622203053295,
                "wh"  : 0.0011622203053295,
                "flb" : 99.2878733152102,
                "BTU" : 0.00396564972437776,
                "btu" : 0.00396564972437776
            ],
            "cal" : [
                "J"   : 4.18679484613929,
                "e"   : 41867928.3372801,
                "c"   : 1.00067013349059,
                "cal" : 1,
                "eV"  : 26131700000000000000,
                "ev"  : 26131700000000000000,
                "HPh" : 0.00000155960800463137,
                "hh"  : 0.00000155960800463137,
                "Wh"  : 0.00116299914807955,
                "wh"  : 0.00116299914807955,
                "flb" : 99.3544094443283,
                "BTU" : 0.00396830723907002,
                "btu" : 0.00396830723907002
            ],
            "eV" : [
                "J"   : 0.000000000000000000160219000146921,
                "e"   : 0.00000000000160218923136574,
                "c"   : 0.0000000000000000000382933423195043,
                "cal" : 0.0000000000000000000382676978535648,
                "eV"  : 1,
                "ev"  : 1,
                "HPh" : 0.0000000000000000000000000596826078912344,
                "hh"  : 0.0000000000000000000000000596826078912344,
                "Wh"  : 0.0000000000000000000000445053000026614,
                "wh"  : 0.0000000000000000000000445053000026614,
                "flb" : 0.00000000000000000380206452103492,
                "BTU" : 0.000000000000000000000151857982414846,
                "btu" : 0.000000000000000000000151857982414846
            ],
            "ev" : [
                "J"   : 0.000000000000000000160219000146921,
                "e"   : 0.000000000000160218923136574,
                "c"   : 0.0000000000000000000382933423195043,
                "cal" : 0.0000000000000000000382676978535648,
                "eV"  : 1,
                "ev"  : 1,
                "HPh" : 0.0000000000000000000000000596826078912344,
                "hh"  : 0.0000000000000000000000000596826078912344,
                "Wh"  : 0.0000000000000000000000445053000026614,
                "wh"  : 0.0000000000000000000000445053000026614,
                "flb" : 0.00000000000000000380206452103492,
                "BTU" : 0.000000000000000000000151857982414846,
                "btu" : 0.000000000000000000000151857982414846
            ],
            "HPh" : [
                "J"   : 2684517.41316170,
                "e"   : 26845161228302.4,
                "c"   : 641616.438565991,
                "cal" : 641186.757845835,
                "eV"  : 16755300000000000000000000,
                "ev"  : 16755300000000000000000000,
                "HPh" : 1,
                "hh"  : 1,
                "Wh"  : 745.699653134593,
                "wh"  : 745.699653134593,
                "flb" : 63704731.6692964,
                "BTU" : 2544.42605275546,
                "btu" : 2544.42605275546
            ],
            "hh" : [
                "J"   : 2684517.4131617,
                "e"   : 26845161228302.4,
                "c"   : 641616.438565991,
                "cal" : 641186.757845835,
                "eV"  : 16755300000000000000000000,
                "ev"  : 16755300000000000000000000,
                "HPh" : 1,
                "hh"  : 1,
                "Wh"  : 745.699653134593,
                "wh"  : 745.699653134593,
                "flb" : 63704731.6692964,
                "BTU" : 2544.42605275546,
                "btu" : 2544.42605275546
            ],
            "Wh" : [
                "J"   : 3599.9982055472,
                "e"   : 35999964751.8369,
                "c"   : 860.422069219046,
                "cal" : 859.845857713046,
                "eV"  : 22469234000000000000000,
                "ev"  : 22469234000000000000000,
                "HPh" : 0.00134102248243839,
                "hh"  : 0.00134102248243839,
                "Wh"  : 1,
                "wh"  : 1,
                "flb" : 85429.4774062316,
                "BTU" : 3.41213254164705,
                "btu" : 3.41213254164705
            ],
            "wh" : [
                "J"   : 3599.9982055472,
                "e"   : 35999964751.8369,
                "c"   : 860.422069219046,
                "cal" : 8590845857713046,
                "eV"  : 22469234000000000000000,
                "ev"  : 22469234000000000000000,
                "HPh" : 0.00134102248243839,
                "hh"  : 0.00134102248243839,
                "Wh"  : 1,
                "wh"  : 1,
                "flb" : 85429.4774062316,
                "BTU" : 3.41213254164705,
                "btu" : 3.41213254164705
            ],
            "flb" : [
                "J"   : 0.0421400003236424,
                "e"   : 421399.80068766,
                "c"   : 0.0100717234301644,
                "cal" : 0.0100649785509554,
                "eV"  : 263015000000000000,
                "ev"  : 263015000000000000,
                "HPh" : 0.000000015697421114513,
                "hh"  : 0.000000015697421114513,
                "Wh"  : 0.0000117055614802,
                "wh"  : 0.0000117055614802,
                "flb" : 1,
                "BTU" : 0.0000399409272448406,
                "btu" : 0.0000399409272448406
            ],
            "BTU" : [
                "J"   : 1055.05813786749,
                "e"   : 10550576307.4665,
                "c"   : 252.165488508168,
                "cal" : 251.99661713551,
                "eV"  : 6585100000000000000000,
                "ev"  : 6585100000000000000000,
                "HPh" : 0.000393015941224568,
                "hh"  : 0.000393015941224568,
                "Wh"  : 0.293071851047526,
                "wh"  : 0.293071851047526,
                "flb" : 25036.9750774671,
                "BTU" : 1,
                "btu" : 1
            ],
            "btu" : [
                "J"   : 1055.05813786749,
                "e"   : 10550576307.4665,
                "c"   : 252.165488508168,
                "cal" : 251.996617135510,
                "eV"  : 6585100000000000000000,
                "ev"  : 6585100000000000000000,
                "HPh" : 0.000393015941224568,
                "hh"  : 0.000393015941224568,
                "Wh"  : 0.293071851047526,
                "wh"  : 0.293071851047526,
                "flb" : 25036.9750774671,
                "BTU" : 1,
                "btu" : 1
            ]
        ],
        "Power" : [
            "HP" : [
                "HP" : 1,
                "h"  : 1,
                "W"  : 745.701,
                "w"  : 745.701
            ],
            "h" : [
                "HP" : 1,
                "h"  : 1,
                "W"  : 745.701,
                "w"  : 745.701
            ],
            "W" : [
                "HP" : 0.00134102006031908,
                "h"  : 0.00134102006031908,
                "W"  : 1,
                "w"  : 1
            ],
            "w" : [
                "HP" : 0.00134102006031908,
                "h"  : 0.00134102006031908,
                "W"  : 1,
                "w"  : 1
            ]
        ],
        "Magnetism" : [
            "T" : [
                "T"  : 1,
                "ga" : 10000
            ],
            "ga" : [
                "T"  : 0.0001,
                "ga" : 1
            ]
        ],
        "Liquid" : [
            "tsp" : [
                "tsp"   : 1,
                "tbs"   : 0.333333333333333,
                "oz"    : 0.166666666666667,
                "cup"   : 0.0208333333333333,
                "pt"    : 0.0104166666666667,
                "us_pt" : 0.0104166666666667,
                "uk_pt" : 0.0086755851682196,
                "qt"    : 0.00520833333333333,
                "gal"   : 0.00130208333333333,
                "l"     : 0.0049299940840071,
                "lt"    : 0.0049299940840071
            ],
            "tbs" : [
                "tsp"   : 3,
                "tbs"   : 1,
                "oz"    : 0.5,
                "cup"   : 0.0625,
                "pt"    : 0.03125,
                "us_pt" : 0.03125,
                "uk_pt" : 0.0260267555046588,
                "qt"    : 0.015625,
                "gal"   : 0.00390625,
                "l"     : 0.0147899822520213,
                "lt"    : 0.0147899822520213
            ],
            "oz" : [
                "tsp"   : 6,
                "tbs"   : 2,
                "oz"    : 1,
                "cup"   : 0.125,
                "pt"    : 0.0625,
                "us_pt" : 0.0625,
                "uk_pt" : 0.0520535110093176,
                "qt"    : 0.03125,
                "gal"   : 0.0078125,
                "l"     : 0.0295799645040426,
                "lt"    : 0.0295799645040426
            ],
            "cup" : [
                "tsp"   : 48,
                "tbs"   : 16,
                "oz"    : 8,
                "cup"   : 1.0,
                "pt"    : 0.5,
                "us_pt" : 0.5,
                "uk_pt" : 0.416428088074541,
                "qt"    : 0.25,
                "gal"   : 0.0625,
                "l"     : 0.236639716032341,
                "lt"    : 0.236639716032341
            ],
            "pt" : [
                "tsp"   : 96,
                "tbs"   : 32,
                "oz"    : 16,
                "cup"   : 2,
                "pt"    : 1.0,
                "us_pt" : 1.0,
                "uk_pt" : 0.832856176149081,
                "qt"    : 0.5,
                "gal"   : 0.125,
                "l"     : 0.473279432064682,
                "lt"    : 0.473279432064682
            ],
            "us_pt" : [
                "tsp"   : 96,
                "tbs"   : 32,
                "oz"    : 16,
                "cup"   : 2,
                "pt"    : 1.0,
                "us_pt" : 1.0,
                "uk_pt" : 0.832856176149081,
                "qt"    : 0.5,
                "gal"   : 0.125,
                "l"     : 0.473279432064682,
                "lt"    : 0.473279432064682
            ],
            "uk_pt" : [
                "tsp"   : 115.266,
                "tbs"   : 38.422,
                "oz"    : 19.211,
                "cup"   : 2.401375,
                "pt"    : 1.2006875,
                "us_pt" : 1.2006875,
                "uk_pt" : 1.0,
                "qt"    : 0.60034375,
                "gal"   : 0.1500859375,
                "l"     : 0.568260698087162,
                "lt"    : 0.568260698087162
            ],
            "qt" : [
                "tsp"   : 192,
                "tbs"   : 64,
                "oz"    : 32,
                "cup"   : 4,
                "pt"    : 2,
                "us_pt" : 2,
                "uk_pt" : 1.66571235229816,
                "qt"    : 1.0,
                "gal"   : 0.25,
                "l"     : 0.946558864129363,
                "lt"    : 0.946558864129363
            ],
            "gal" : [
                "tsp"   : 768,
                "tbs"   : 256,
                "oz"    : 128,
                "cup"   : 16,
                "pt"    : 8,
                "us_pt" : 8,
                "uk_pt" : 6.66284940919265,
                "qt"    : 4,
                "gal"   : 1.0,
                "l"     : 3.78623545651745,
                "lt"    : 3.78623545651745
            ],
            "l" : [
                "tsp"   : 202.84,
                "tbs"   : 67.6133333333333,
                "oz"    : 33.8066666666667,
                "cup"   : 4.22583333333333,
                "pt"    : 2.11291666666667,
                "us_pt" : 2.11291666666667,
                "uk_pt" : 1.75975569552166,
                "qt"    : 1.05645833333333,
                "gal"   : 0.264114583333333,
                "l"     : 1.0,
                "lt"    : 1.0
            ],
            "lt" : [
                "tsp"   : 202.84,
                "tbs"   : 67.6133333333333,
                "oz"    : 33.8066666666667,
                "cup"   : 4.22583333333333,
                "pt"    : 2.11291666666667,
                "us_pt" : 2.11291666666667,
                "uk_pt" : 1.75975569552166,
                "qt"    : 1.05645833333333,
                "gal"   : 0.264114583333333,
                "l"     : 1.0,
                "lt"    : 1.0
            ]
        ]
    ];


    /**
     * parseComplex
     *
     * Parses a complex number into its real and imaginary parts, and an I or J suffix
     *
     * @param    string        complexNumber    The complex number
     * @return    string[]    Indexed on "real", "imaginary" and "suffix"
     */
    public static function parseComplex(var complexNumber)
    {
        var realNumber, imaginary, suffix, leadingSign, power;
        string workString;
        
        let workString = (string) complexNumber;

        let realNumber = 0;
        let imaginary = 0;
        //    Extract the suffix, if there is one
        let suffix = substr(workString, -1);
        if (!is_numeric(suffix)) {
            let workString = substr(workString, 0, -1);
        } else {
            let suffix = "";
        }

        //    Split the input into its Real and Imaginary components
        let leadingSign = 0;
        
        if (strlen(workString) > 0 && ((substr(workString, 0, 1) == "+") || (substr(workString, 0, 1) == "-"))) {
            let leadingSign = 1;
        }
        
        let power = "";
        let realNumber = call_user_func("strtok", workString, "+-");
        
        if (strtoupper(substr(realNumber, -1)) == "E") {
            let power = call_user_func("strtok", "+-");
            let leadingSign = leadingSign + 1;
        }

        let realNumber = substr(workString, 0, strlen(realNumber) + strlen(power) + leadingSign);
        
        if (suffix != "") {
            let imaginary = substr(workString, strlen(realNumber));

            if (empty(imaginary) && (empty(realNumber) || (realNumber == "+") || (realNumber == "-"))) {
                let imaginary = realNumber . "1";
                let realNumber = "0";
            } elseif (empty(imaginary)) {
                let imaginary = realNumber;
                let realNumber = "0";
            } elseif ((imaginary == "+") || (imaginary == "-")) {
                let imaginary = imaginary . "1";
            }
        }

        return [
            "real": realNumber,
            "imaginary": imaginary,
            "suffix": suffix
        ];
    }


    /**
     * Cleans the leading characters in a complex number string
     *
     * @param    string        complexNumber    The complex number to clean
     * @return    string        The "cleaned" complex number
     */
    private static function cleanComplex(string complexNumber)
    {
        var firstLetter = substr(complexNumber, 0, 1);
        
        if (firstLetter == "+") {
            let complexNumber = substr(complexNumber, 1);
        }
        
        if (firstLetter == "\0") {
            let complexNumber = substr(complexNumber, 1);
        }
        
        if (firstLetter == ".") {
            let complexNumber = "0" . complexNumber;
        }
        
        if (firstLetter == "+") {
            let complexNumber = substr(complexNumber, 1);
        }
        
        return complexNumber;
    }

    /**
     * Formats a number base string value with leading zeroes
     *
     * @param    string        xVal        The "number" to pad
     * @param    integer        places        The length that we want to pad this value
     * @return    string        The padded "number"
     */
    private static function nbrConversionFormat(xVal, places)
    {
        if (!is_null(places)) {
            if (strlen(xVal) <= places) {
                return substr(str_pad(xVal, places, "0", STR_PAD_LEFT), -10);
            } else {
                return \ZExcel\Calculation\Functions::NaN();
            }
        }

        return substr(xVal, -10);
    }

    /**
     *    BESSELI
     *
     *    Returns the modified Bessel function In(x), which is equivalent to the Bessel function evaluated
     *        for purely imaginary arguments
     *
     *    Excel Function:
     *        BESSELI(x,ord)
     *
     *    @access    public
     *    @category Engineering Functions
     *    @param    float        x        The value at which to evaluate the function.
     *                                If x is nonnumeric, BESSELI returns the #VALUE! error value.
     *    @param    integer        ord    The order of the Bessel function.
     *                                If ord is not an integer, it is truncated.
     *                                If ord is nonnumeric, BESSELI returns the #VALUE! error value.
     *                                If ord < 0, BESSELI returns the #NUM! error value.
     *    @return    float
     *
     */
    public static function besseli(var x, var ord)
    {
        var fSqrX, f_2_PI, fXAbs;
        int ordK;
        double fResult, fTerm;
        
        let x   = (is_null(x))   ? 0.0 : \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let ord = (is_null(ord)) ? 0.0 : \ZExcel\Calculation\Functions::flattenSingleValue(ord);

        if ((is_numeric(x)) && (is_numeric(ord))) {
            let ord = floor(ord);
            if (ord < 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }

            if (abs(x) <= 30) {
                let fResult = pow(x / 2, ord) / \ZExcel\Calculation\MathTrig::FaCT(ord);
                let fTerm = fResult;
                let ordK = 1;
                let fSqrX = (x * x) / 4;
                do {
                    let fTerm = fTerm * fSqrX;
                    let fTerm = fTerm / (ordK * (ordK + ord));
                    let fResult = fResult + fTerm;
                    let ordK = ordK + 1;
                } while ((abs(fTerm) > 0.0000000000001) && (ordK < 100));
            } else {
                let f_2_PI = 2 * M_PI;

                let fXAbs = abs(x);
                let fResult = exp(fXAbs) / sqrt(f_2_PI * fXAbs);
                if ((ord & 1) && (x < 0)) {
                    let fResult = -fResult;
                }
            }
            
            if (!is_nan(fResult)) {
                return fResult;
            } else {
                return \ZExcel\Calculation\Functions::NaN();
            }
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     *    BESSELJ
     *
     *    Returns the Bessel function
     *
     *    Excel Function:
     *        BESSELJ(x,ord)
     *
     *    @access    public
     *    @category Engineering Functions
     *    @param    float        x        The value at which to evaluate the function.
     *                                If x is nonnumeric, BESSELJ returns the #VALUE! error value.
     *    @param    integer        ord    The order of the Bessel function. If n is not an integer, it is truncated.
     *                                If ord is nonnumeric, BESSELJ returns the #VALUE! error value.
     *                                If ord < 0, BESSELJ returns the #NUM! error value.
     *    @return    float
     *
     */
    public static function besselj(x, ord)
    {
        var fXAbs;
        double fResult, fTerm, ordK, fSqrX, f_PI_DIV_2, f_PI_DIV_4;
        
        let x   = (is_null(x))   ? 0.0 : \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let ord = (is_null(ord)) ? 0.0 : \ZExcel\Calculation\Functions::flattenSingleValue(ord);

        if ((is_numeric(x)) && (is_numeric(ord))) {
            let ord = floor(ord);
            
            if (ord < 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }

            let fResult = 0;
            if (abs(x) <= 30) {
                let fResult = pow(x / 2, ord) / \ZExcel\Calculation\MathTrig::FaCT(ord);
                let fTerm = fResult;
                
                let ordK = 1;
                let fSqrX = (x * x) / -4;
                do {
                    let fTerm = fTerm * fSqrX;
                    let fTerm = fTerm / (ordK * (ordK + ord));
                    let fResult = fResult + fTerm;
                    let ordK = ordK + 1;
                } while ((abs(fTerm) > 0.000000000001) && (ordK < 100));
            } else {
                let f_PI_DIV_2 = M_PI / 2;
                let f_PI_DIV_4 = M_PI / 4;

                let fXAbs = abs(x);
                let fResult = sqrt(\ZExcel\Calculation\Functions::M_2DIVPI / (double) fXAbs) * cos((double) fXAbs - (double) ord * f_PI_DIV_2 - f_PI_DIV_4);
                if ((ord & 1) && (x < 0)) {
                    let fResult = -fResult;
                }
            }
            return (is_nan(fResult)) ? \ZExcel\Calculation\Functions::NaN() : fResult;
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    private static function besselK0(double fNum) -> double
    {
        double y, fRet, fNum2;
        
        if (fNum <= 2) {
            let fNum2 = fNum * 0.5;
            let y = (fNum2 * fNum2);
            let fRet = -log(fNum2) * self::BeSSELI(fNum, 0) + (-0.57721566 + y * (0.42278420 + y * (0.23069756 + y * (0.0348859 + y * (0.00262698 + y * (0.0001075 + y * 0.0000074))))));
        } else {
            let y = 2 / fNum;
            let fRet = exp(-fNum) / sqrt(fNum) * (1.25331414 + y * (-0.07832358 + y * (0.02189568 + y * (-0.01062446 + y * (0.00587872 + y * (-0.00251540 + y * 0.00053208))))));
        }
        return fRet;
    }


    private static function besselK1(double fNum) -> double
    {
        double y, fRet, fNum2;
        
        if (fNum <= 2) {
            let fNum2 = fNum * 0.5;
            let y = (fNum2 * fNum2);
            let fRet = log(fNum2) * self::BeSSELI(fNum, 1) + (1 + y * (0.15443144 + y * (-0.67278579 + y * (-0.18156897 + y * (-0.01919402 + y * (-0.00110404 + y * (-0.00004686))))))) / fNum;
        } else {
            let y = 2 / fNum;
            let fRet = exp(-fNum) / sqrt(fNum) * (1.25331414 + y * (0.23498619 + y * (-0.0365562 + y * (0.01504268 + y * (-0.00780353 + y * (0.00325614 + y * (-0.00068245)))))));
        }
        
        return fRet;
    }


    /**
     *    BESSELK
     *
     *    Returns the modified Bessel function Kn(x), which is equivalent to the Bessel functions evaluated
     *        for purely imaginary arguments.
     *
     *    Excel Function:
     *        BESSELK(x,ord)
     *
     *    @access    public
     *    @category Engineering Functions
     *    @param    float        x        The value at which to evaluate the function.
     *                                If x is nonnumeric, BESSELK returns the #VALUE! error value.
     *    @param    integer        ord    The order of the Bessel function. If n is not an integer, it is truncated.
     *                                If ord is nonnumeric, BESSELK returns the #VALUE! error value.
     *                                If ord < 0, BESSELK returns the #NUM! error value.
     *    @return    float
     *
     */
    public static function besselk(var x, var ord)
    {
        double fBkm, fBkp, fBk, fTox;
        int n;
        
        let x   = (is_null(x))   ? 0.0 : \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let ord = (is_null(ord)) ? 0.0 : \ZExcel\Calculation\Functions::flattenSingleValue(ord);

        if ((is_numeric(x)) && (is_numeric(ord))) {
            if ((ord < 0) || (x == 0.0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }

            switch (floor(ord)) {
                case 0:
                    return self::besselK0(x);
                case 1:
                    return self::besselK1(x);
                default:
                    let fTox = 2 / x;
                    let fBkm = 0 + self::besselK0(x);
                    let fBk  = 0 + self::besselK1(x);
                    
                    for n in range(1, ord - 1) {
                        let fBkp = fBkm + (n * fTox * fBk);
                        let fBkm = fBk;
                        let fBk  = fBkp;
                    }
            }
            
            if (!is_nan(fBk)) {
                return fBk;
            } else {
                return \ZExcel\Calculation\Functions::NaN();
            }
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    private static function besselY0(fNum)
    {
        var y, z, xx, f1, f2, fRet;
        
        if (fNum < 8.0) {
            let y = (fNum * fNum);
            let f1 = -2957821389.0 + y * (7062834065.0 + y * (-512359803.6 + y * (10879881.29 + y * (-86327.92757 + y * 228.4622733))));
            let f2 = 40076544269.0 + y * (745249964.8 + y * (7189466.438 + y * (47447.26470 + y * (226.1030244 + y))));
            let fRet = f1 / f2 + 0.636619772 * self::besselj(fNum, 0) * log(fNum);
        } else {
            let z = 8.0 / fNum;
            let y = (z * z);
            let xx = fNum - 0.785398164;
            let f1 = 1 + y * (-0.001098628627 + y * (0.00002734510407 + y * (-0.000002073370639 + y * 0.0000002093887211)));
            let f2 = -0.01562499995 + y * (0.0001430488765 + y * (-0.000006911147651 + y * (0.0000007621095161 + y * (-0.0000000934945152))));
            let fRet = sqrt(0.636619772 / fNum) * (sin(xx) * f1 + z * cos(xx) * f2);
        }
        
        return fRet;
    }


    private static function besselY1(fNum)
    {
        var y, f1, f2, fRet;
        
        if (fNum < 8.0) {
            let y = (fNum * fNum);
            let f1 = fNum * (-4900604943000 + y * (1275274390000 + y * (-51534381390 + y * (734926455.1 + y * (-4237922.726 + y * 8511.937935)))));
            let f2 = 24995805700000 + y * (424441966400 + y * (3733650367 + y * (22459040.02 + y * (102042.605 + y * (354.9632885 + y)))));
            let fRet = f1 / f2 + 0.636619772 * ( self::besselj(fNum, 1) * log(fNum) - 1 / fNum);
        } else {
            let fRet = sqrt(0.636619772 / fNum) * sin(fNum - 2.356194491);
        }
        
        return fRet;
    }


    /**
     *    BESSELY
     *
     *    Returns the Bessel function, which is also called the Weber function or the Neumann function.
     *
     *    Excel Function:
     *        BESSELY(x,ord)
     *
     *    @access    public
     *    @category Engineering Functions
     *    @param    float        x        The value at which to evaluate the function.
     *                                If x is nonnumeric, BESSELK returns the #VALUE! error value.
     *    @param    integer        ord    The order of the Bessel function. If n is not an integer, it is truncated.
     *                                If ord is nonnumeric, BESSELK returns the #VALUE! error value.
     *                                If ord < 0, BESSELK returns the #NUM! error value.
     *
     *    @return    float
     */
    public static function bessely(x, ord)
    {
        var fBym, fBy;
        double fTox, fByp;
        int n;
        
        let x   = (is_null(x)) ? 0.0 : \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let ord = (is_null(ord)) ? 0.0 : \ZExcel\Calculation\Functions::flattenSingleValue(ord);

        if ((is_numeric(x)) && (is_numeric(ord))) {
            if ((ord < 0) || (x == 0.0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }

            switch(floor(ord)) {
                case 0:
                    return self::besselY0(x);
                case 1:
                    return self::besselY1(x);
                default:
                    let fTox = 2 / x;
                    let fBym = self::besselY0(x);
                    let fBy  = self::besselY1(x);
                    for n in range(1, ord - 1) {
                        let fByp = n * fTox * fBy - fBym;
                        let fBym = fBy;
                        let fBy  = fByp;
                    }
            }
            return (is_nan(fBy)) ? \ZExcel\Calculation\Functions::NaN() : fBy;
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * BINTODEC
     *
     * Return a binary value as decimal.
     *
     * Excel Function:
     *        BIN2DEC(x)
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        x        The binary number (as a string) that you want to convert. The number
     *                                cannot contain more than 10 characters (10 bits). The most significant
     *                                bit of number is the sign bit. The remaining 9 bits are magnitude bits.
     *                                Negative numbers are represented using two"s-complement notation.
     *                                If number is not a valid binary number, or if number contains more than
     *                                10 characters (10 bits), BIN2DEC returns the #NUM! error value.
     * @return    string
     */
    public static function bintodec(x)
    {
        array out = [];
        
        let x = \ZExcel\Calculation\Functions::flattenSingleValue(x);

        if (is_bool(x)) {
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
                let x = (int) x;
            } else {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
        }
        if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC) {
            let x = floor(x);
        }
        
        let x = (string) x;
        
        if (strlen(x) > preg_match_all("/[01]/", x, out)) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        if (strlen(x) > 10) {
            return \ZExcel\Calculation\Functions::NaN();
        } elseif (strlen(x) == 10) {
            //    Two"s Complement
            let x = substr(x, -9);
            return "-" . (512-bindec(x));
        }
        return bindec(x);
    }


    /**
     * BINTOHEX
     *
     * Return a binary value as hex.
     *
     * Excel Function:
     *        BIN2HEX(x[,places])
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        x        The binary number (as a string) that you want to convert. The number
     *                                cannot contain more than 10 characters (10 bits). The most significant
     *                                bit of number is the sign bit. The remaining 9 bits are magnitude bits.
     *                                Negative numbers are represented using two"s-complement notation.
     *                                If number is not a valid binary number, or if number contains more than
     *                                10 characters (10 bits), BIN2HEX returns the #NUM! error value.
     * @param    integer        places    The number of characters to use. If places is omitted, BIN2HEX uses the
     *                                minimum number of characters necessary. Places is useful for padding the
     *                                return value with leading 0s (zeros).
     *                                If places is not an integer, it is truncated.
     *                                If places is nonnumeric, BIN2HEX returns the #VALUE! error value.
     *                                If places is negative, BIN2HEX returns the #NUM! error value.
     * @return    string
     */
    public static function bintohex(var x, var places = null)
    {
        var hexVal;
        
        let x = \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let places = \ZExcel\Calculation\Functions::flattenSingleValue(places);

        if (is_bool(x)) {
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
                let x = (int) x;
            } else {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
        }
        
        if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC) {
            let x = floor(x);
        }
        
        let x = (string) x;
        
        if (strlen(x) > preg_match_all("/[01]/", x)) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        if (strlen(x) > 10) {
            return \ZExcel\Calculation\Functions::NaN();
        } elseif (strlen(x) == 10) {
            //    Two"s Complement
            return str_repeat("F", 8).substr(strtoupper(dechex(bindec(substr(x, -9)))), -2);
        }
        
        let hexVal = (string) strtoupper(dechex(bindec(x)));

        return self::nbrConversionFormat(hexVal, places);
    }


    /**
     * BINTOOCT
     *
     * Return a binary value as octal.
     *
     * Excel Function:
     *        BIN2OCT(x[,places])
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        x        The binary number (as a string) that you want to convert. The number
     *                                cannot contain more than 10 characters (10 bits). The most significant
     *                                bit of number is the sign bit. The remaining 9 bits are magnitude bits.
     *                                Negative numbers are represented using two"s-complement notation.
     *                                If number is not a valid binary number, or if number contains more than
     *                                10 characters (10 bits), BIN2OCT returns the #NUM! error value.
     * @param    integer        places    The number of characters to use. If places is omitted, BIN2OCT uses the
     *                                minimum number of characters necessary. Places is useful for padding the
     *                                return value with leading 0s (zeros).
     *                                If places is not an integer, it is truncated.
     *                                If places is nonnumeric, BIN2OCT returns the #VALUE! error value.
     *                                If places is negative, BIN2OCT returns the #NUM! error value.
     * @return    string
     */
    public static function bintooct(var x, var places = null)
    {
        var octVal;
        array out = [];
        
        let x = \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let places = \ZExcel\Calculation\Functions::flattenSingleValue(places);

        if (is_bool(x)) {
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
                let x = (int) x;
            } else {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
        }
        
        if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC) {
            let x = floor(x);
        }
        
        let x = (string) x;
        
        if (strlen(x) > preg_match_all("/[01]/", x, out)) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        if (strlen(x) > 10) {
            return \ZExcel\Calculation\Functions::NaN();
        } elseif (strlen(x) == 10) {
            //    Two"s Complement
            return str_repeat("7", 7).substr(strtoupper(decoct(bindec(substr(x, -9)))), -3);
        }
        
        let octVal = (string) decoct(bindec(x));

        return self::nbrConversionFormat(octVal, places);
    }


    /**
     * DECTOBIN
     *
     * Return a decimal value as binary.
     *
     * Excel Function:
     *        DEC2BIN(x[,places])
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        x        The decimal integer you want to convert. If number is negative,
     *                                valid place values are ignored and DEC2BIN returns a 10-character
     *                                (10-bit) binary number in which the most significant bit is the sign
     *                                bit. The remaining 9 bits are magnitude bits. Negative numbers are
     *                                represented using two"s-complement notation.
     *                                If number < -512 or if number > 511, DEC2BIN returns the #NUM! error
     *                                value.
     *                                If number is nonnumeric, DEC2BIN returns the #VALUE! error value.
     *                                If DEC2BIN requires more than places characters, it returns the #NUM!
     *                                error value.
     * @param    integer        places    The number of characters to use. If places is omitted, DEC2BIN uses
     *                                the minimum number of characters necessary. Places is useful for
     *                                padding the return value with leading 0s (zeros).
     *                                If places is not an integer, it is truncated.
     *                                If places is nonnumeric, DEC2BIN returns the #VALUE! error value.
     *                                If places is zero or negative, DEC2BIN returns the #NUM! error value.
     * @return    string
     */
    public static function dectobin(var x, var places = null)
    {
        var r;
        
        let x      = \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let places = \ZExcel\Calculation\Functions::flattenSingleValue(places);

        if (is_bool(x)) {
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
                let x = (int) x;
            } else {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
        }
        
        let x = (string) x;
        
        if (strlen(x) > preg_match_all("/[-0123456789.]/", x)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let x = strval(floor(x));
        let r = decbin(x);
        
        if (strlen(r) == 32) {
            //    Two's Complement
            let r = substr(r, -10);
        } elseif (strlen(r) > 11) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        return self::nbrConversionFormat(r, places);
    }


    /**
     * DECTOHEX
     *
     * Return a decimal value as hex.
     *
     * Excel Function:
     *        DEC2HEX(x[,places])
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        x        The decimal integer you want to convert. If number is negative,
     *                                places is ignored and DEC2HEX returns a 10-character (40-bit)
     *                                hexadecimal number in which the most significant bit is the sign
     *                                bit. The remaining 39 bits are magnitude bits. Negative numbers
     *                                are represented using two"s-complement notation.
     *                                If number < -549,755,813,888 or if number > 549,755,813,887,
     *                                DEC2HEX returns the #NUM! error value.
     *                                If number is nonnumeric, DEC2HEX returns the #VALUE! error value.
     *                                If DEC2HEX requires more than places characters, it returns the
     *                                #NUM! error value.
     * @param    integer        places    The number of characters to use. If places is omitted, DEC2HEX uses
     *                                the minimum number of characters necessary. Places is useful for
     *                                padding the return value with leading 0s (zeros).
     *                                If places is not an integer, it is truncated.
     *                                If places is nonnumeric, DEC2HEX returns the #VALUE! error value.
     *                                If places is zero or negative, DEC2HEX returns the #NUM! error value.
     * @return    string
     */
    public static function dectohex(var x, var places = null)
    {
        var r;
        
        let x      = \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let places = \ZExcel\Calculation\Functions::flattenSingleValue(places);

        if (is_bool(x)) {
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
                let x = (int) x;
            } else {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
        }
        
        let x = (string) x;
        
        if (strlen(x) > preg_match_all("/[-0123456789.]/", x)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let x = strval(floor(x));
        let r = strtoupper(dechex(x));
        
        if (strlen(r) == 8) {
            //    Two's Complement
            let r = "FF" . r;
        }

        return self::nbrConversionFormat(r, places);
    }


    /**
     * DECTOOCT
     *
     * Return an decimal value as octal.
     *
     * Excel Function:
     *        DEC2OCT(x[,places])
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        x        The decimal integer you want to convert. If number is negative,
     *                                places is ignored and DEC2OCT returns a 10-character (30-bit)
     *                                octal number in which the most significant bit is the sign bit.
     *                                The remaining 29 bits are magnitude bits. Negative numbers are
     *                                represented using two"s-complement notation.
     *                                If number < -536,870,912 or if number > 536,870,911, DEC2OCT
     *                                returns the #NUM! error value.
     *                                If number is nonnumeric, DEC2OCT returns the #VALUE! error value.
     *                                If DEC2OCT requires more than places characters, it returns the
     *                                #NUM! error value.
     * @param    integer        places    The number of characters to use. If places is omitted, DEC2OCT uses
     *                                the minimum number of characters necessary. Places is useful for
     *                                padding the return value with leading 0s (zeros).
     *                                If places is not an integer, it is truncated.
     *                                If places is nonnumeric, DEC2OCT returns the #VALUE! error value.
     *                                If places is zero or negative, DEC2OCT returns the #NUM! error value.
     * @return    string
     */
    public static function dectooct(var x, var places = null)
    {
        var r;
        
        let x      = \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let places = \ZExcel\Calculation\Functions::flattenSingleValue(places);

        if (is_bool(x)) {
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
                let x = (int) x;
            } else {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
        }
        
        let x = (string) x;
        
        if (strlen(x) > preg_match_all("/[-0123456789.]/", x)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let x = strval(floor(x));
        let r = decoct(x);
        if (strlen(r) == 11) {
            //    Two's Complement
            let r = substr(r, -10);
        }

        return self::nbrConversionFormat(r, places);
    }


    /**
     * HEXTOBIN
     *
     * Return a hex value as binary.
     *
     * Excel Function:
     *        HEX2BIN(x[,places])
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        x            the hexadecimal number you want to convert. Number cannot
     *                                    contain more than 10 characters. The most significant bit of
     *                                    number is the sign bit (40th bit from the right). The remaining
     *                                    9 bits are magnitude bits. Negative numbers are represented
     *                                    using two"s-complement notation.
     *                                    If number is negative, HEX2BIN ignores places and returns a
     *                                    10-character binary number.
     *                                    If number is negative, it cannot be less than FFFFFFFE00, and
     *                                    if number is positive, it cannot be greater than 1FF.
     *                                    If number is not a valid hexadecimal number, HEX2BIN returns
     *                                    the #NUM! error value.
     *                                    If HEX2BIN requires more than places characters, it returns
     *                                    the #NUM! error value.
     * @param    integer        places        The number of characters to use. If places is omitted,
     *                                    HEX2BIN uses the minimum number of characters necessary. Places
     *                                    is useful for padding the return value with leading 0s (zeros).
     *                                    If places is not an integer, it is truncated.
     *                                    If places is nonnumeric, HEX2BIN returns the #VALUE! error value.
     *                                    If places is negative, HEX2BIN returns the #NUM! error value.
     * @return    string
     */
    public static function hextobin(x, places = null)
    {
        var binVal;
        
        let x      = \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let places = \ZExcel\Calculation\Functions::flattenSingleValue(places);

        if (is_bool(x)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let x = (string) x;
        
        if (strlen(x) > preg_match_all("/[0123456789ABCDEF]/", strtoupper(x))) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        let binVal = decbin(hexdec(x));

        return substr(self::nbrConversionFormat(binVal, places), -10);
    }


    /**
     * HEXTODEC
     *
     * Return a hex value as decimal.
     *
     * Excel Function:
     *        HEX2DEC(x)
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        x        The hexadecimal number you want to convert. This number cannot
     *                                contain more than 10 characters (40 bits). The most significant
     *                                bit of number is the sign bit. The remaining 39 bits are magnitude
     *                                bits. Negative numbers are represented using two"s-complement
     *                                notation.
     *                                If number is not a valid hexadecimal number, HEX2DEC returns the
     *                                #NUM! error value.
     * @return    string
     */
    public static function hexToDec(var x)
    {
        let x = \ZExcel\Calculation\Functions::flattenSingleValue(x);

        if (is_bool(x)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let x = (string) x;
        
        if (strlen(x) > preg_match_all("/[0123456789ABCDEF]/", strtoupper(x))) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        return hexdec(x);
    }


    /**
     * HEXTOOCT
     *
     * Return a hex value as octal.
     *
     * Excel Function:
     *        HEX2OCT(x[,places])
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        x            The hexadecimal number you want to convert. Number cannot
     *                                    contain more than 10 characters. The most significant bit of
     *                                    number is the sign bit. The remaining 39 bits are magnitude
     *                                    bits. Negative numbers are represented using two"s-complement
     *                                    notation.
     *                                    If number is negative, HEX2OCT ignores places and returns a
     *                                    10-character octal number.
     *                                    If number is negative, it cannot be less than FFE0000000, and
     *                                    if number is positive, it cannot be greater than 1FFFFFFF.
     *                                    If number is not a valid hexadecimal number, HEX2OCT returns
     *                                    the #NUM! error value.
     *                                    If HEX2OCT requires more than places characters, it returns
     *                                    the #NUM! error value.
     * @param    integer        places        The number of characters to use. If places is omitted, HEX2OCT
     *                                    uses the minimum number of characters necessary. Places is
     *                                    useful for padding the return value with leading 0s (zeros).
     *                                    If places is not an integer, it is truncated.
     *                                    If places is nonnumeric, HEX2OCT returns the #VALUE! error
     *                                    value.
     *                                    If places is negative, HEX2OCT returns the #NUM! error value.
     * @return    string
     */
    public static function hextooct(var x, var places = null)
    {
        var octVal;
        
        let x      = \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let places = \ZExcel\Calculation\Functions::flattenSingleValue(places);

        if (is_bool(x)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let x = (string) x;
        
        if (strlen(x) > preg_match_all("/[0123456789ABCDEF]/", strtoupper(x))) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        let octVal = decoct(hexdec(x));

        return self::nbrConversionFormat(octVal, places);
    }    //    function HEXTOOCT()


    /**
     * OCTTOBIN
     *
     * Return an octal value as binary.
     *
     * Excel Function:
     *        OCT2BIN(x[,places])
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        x            The octal number you want to convert. Number may not
     *                                    contain more than 10 characters. The most significant
     *                                    bit of number is the sign bit. The remaining 29 bits
     *                                    are magnitude bits. Negative numbers are represented
     *                                    using two"s-complement notation.
     *                                    If number is negative, OCT2BIN ignores places and returns
     *                                    a 10-character binary number.
     *                                    If number is negative, it cannot be less than 7777777000,
     *                                    and if number is positive, it cannot be greater than 777.
     *                                    If number is not a valid octal number, OCT2BIN returns
     *                                    the #NUM! error value.
     *                                    If OCT2BIN requires more than places characters, it
     *                                    returns the #NUM! error value.
     * @param    integer        places        The number of characters to use. If places is omitted,
     *                                    OCT2BIN uses the minimum number of characters necessary.
     *                                    Places is useful for padding the return value with
     *                                    leading 0s (zeros).
     *                                    If places is not an integer, it is truncated.
     *                                    If places is nonnumeric, OCT2BIN returns the #VALUE!
     *                                    error value.
     *                                    If places is negative, OCT2BIN returns the #NUM! error
     *                                    value.
     * @return    string
     */
    public static function octtobin(var x, var places = null)
    {
        var r;
        
        let x      = \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let places = \ZExcel\Calculation\Functions::flattenSingleValue(places);

        if (is_bool(x)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let x = (string) x;
        
        if (preg_match_all("/[01234567]/", x) != strlen(x)) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        let r = decbin(octdec(x));

        return self::nbrConversionFormat(r, places);
    }


    /**
     * OCTTODEC
     *
     * Return an octal value as decimal.
     *
     * Excel Function:
     *        OCT2DEC(x)
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        x        The octal number you want to convert. Number may not contain
     *                                more than 10 octal characters (30 bits). The most significant
     *                                bit of number is the sign bit. The remaining 29 bits are
     *                                magnitude bits. Negative numbers are represented using
     *                                two"s-complement notation.
     *                                If number is not a valid octal number, OCT2DEC returns the
     *                                #NUM! error value.
     * @return    string
     */
    public static function octtodec(x)
    {
        let x = \ZExcel\Calculation\Functions::flattenSingleValue(x);

        if (is_bool(x)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let x = (string) x;
        
        if (preg_match_all("/[01234567]/", x) != strlen(x)) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        return octdec(x);
    }


    /**
     * OCTTOHEX
     *
     * Return an octal value as hex.
     *
     * Excel Function:
     *        OCT2HEX(x[,places])
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        x            The octal number you want to convert. Number may not contain
     *                                    more than 10 octal characters (30 bits). The most significant
     *                                    bit of number is the sign bit. The remaining 29 bits are
     *                                    magnitude bits. Negative numbers are represented using
     *                                    two"s-complement notation.
     *                                    If number is negative, OCT2HEX ignores places and returns a
     *                                    10-character hexadecimal number.
     *                                    If number is not a valid octal number, OCT2HEX returns the
     *                                    #NUM! error value.
     *                                    If OCT2HEX requires more than places characters, it returns
     *                                    the #NUM! error value.
     * @param    integer        places        The number of characters to use. If places is omitted, OCT2HEX
     *                                    uses the minimum number of characters necessary. Places is useful
     *                                    for padding the return value with leading 0s (zeros).
     *                                    If places is not an integer, it is truncated.
     *                                    If places is nonnumeric, OCT2HEX returns the #VALUE! error value.
     *                                    If places is negative, OCT2HEX returns the #NUM! error value.
     * @return    string
     */
    public static function octToHex(var x, var places = null)
    {
        var hexVal;
        
        let x      = \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let places = \ZExcel\Calculation\Functions::flattenSingleValue(places);

        if (is_bool(x)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let x = (string) x;
        
        if (preg_match_all("/[01234567]/", x) != strlen(x)) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        let hexVal = strtoupper(dechex(octdec(x)));

        return self::nbrConversionFormat(hexVal, places);
    }


    /**
     * COMPLEX
     *
     * Converts real and imaginary coefficients into a complex number of the form x + yi or x + yj.
     *
     * Excel Function:
     *        COMPLEX(realNumber,imaginary[,places])
     *
     * @access    public
     * @category Engineering Functions
     * @param    float        realNumber        The real coefficient of the complex number.
     * @param    float        imaginary        The imaginary coefficient of the complex number.
     * @param    string        suffix            The suffix for the imaginary component of the complex number.
     *                                        If omitted, the suffix is assumed to be "i".
     * @return    string
     */
    public static function complex(var realNumber = 0.0, var imaginary = 0.0, var suffix = "i")
    {
        let realNumber = (is_null(realNumber)) ? 0.0 : \ZExcel\Calculation\Functions::flattenSingleValue(realNumber);
        let imaginary  = (is_null(imaginary))  ? 0.0 : \ZExcel\Calculation\Functions::flattenSingleValue(imaginary);
        let suffix     = (is_null(suffix))     ? "i" : \ZExcel\Calculation\Functions::flattenSingleValue(suffix);

        if (((is_numeric(realNumber)) && (is_numeric(imaginary))) && ((suffix == "i") || (suffix == "j") || (suffix == ""))) {
            let realNumber   = (float) realNumber;
            let imaginary    = (float) imaginary;

            if (suffix == "") {
                let suffix = "i";
            }
            
            if (realNumber == 0.0) {
                if ((imaginary) == 0.0) {
                    return "0";
                } elseif ((imaginary) == 1.0) {
                    return (suffix);
                } elseif ((imaginary) == -1.0) {
                    return ("-" . suffix);
                }
                
                return (strval(imaginary) . suffix);
            } elseif (imaginary == 0.0) {
                return (strval(realNumber));
            } elseif (imaginary == 1.0) {
                return (strval(realNumber) . "+" . suffix);
            } elseif (imaginary == -1.0) {
                return (strval(realNumber) . "-" . suffix);
            }
            
            if (imaginary > 0.0) {
                let imaginary = "+" . strval(imaginary);
            }
            
            return strval(realNumber) . strval(imaginary) . suffix;
        }

        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * IMAGINARY
     *
     * Returns the imaginary coefficient of a complex number in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMAGINARY(complexNumber)
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        complexNumber    The complex number for which you want the imaginary
     *                                         coefficient.
     * @return    float
     */
    public static function imaginary(var complexNumber)
    {
        var parsedComplex;
        
        let complexNumber = \ZExcel\Calculation\Functions::flattenSingleValue(complexNumber);

        let parsedComplex = self::parseComplex(complexNumber);
        
        return parsedComplex["imaginary"];
    }


    /**
     * IMREAL
     *
     * Returns the real coefficient of a complex number in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMREAL(complexNumber)
     *
     * @access    public
     * @category Engineering Functions
     * @param    string        complexNumber    The complex number for which you want the real coefficient.
     * @return    float
     */
    public static function imreal(var complexNumber)
    {
        var parsedComplex;
        
        let complexNumber = \ZExcel\Calculation\Functions::flattenSingleValue(complexNumber);

        let parsedComplex = self::parseComplex(complexNumber);
        
        return parsedComplex["real"];
    }


    /**
     * IMABS
     *
     * Returns the absolute value (modulus) of a complex number in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMABS(complexNumber)
     *
     * @param    string        complexNumber    The complex number for which you want the absolute value.
     * @return    float
     */
    public static function imabs(var complexNumber) -> float
    {
        var parsedComplex;
        
        let complexNumber = \ZExcel\Calculation\Functions::flattenSingleValue(complexNumber);

        let parsedComplex = self::parseComplex(complexNumber);

        return sqrt((parsedComplex["real"] * parsedComplex["real"]) + (parsedComplex["imaginary"] * parsedComplex["imaginary"]));
    }


    /**
     * IMARGUMENT
     *
     * Returns the argument theta of a complex number, i.e. the angle in radians from the real
     * axis to the representation of the number in polar coordinates.
     *
     * Excel Function:
     *        IMARGUMENT(complexNumber)
     *
     * @param    string        complexNumber    The complex number for which you want the argument theta.
     * @return    float
     */
    public static function imargument(complexNumber)
    {
        var parsedComplex;
        
        let complexNumber = \ZExcel\Calculation\Functions::flattenSingleValue(complexNumber);

        let parsedComplex = self::parseComplex(complexNumber);

        if (parsedComplex["real"] == 0.0) {
            if (parsedComplex["imaginary"] == 0.0) {
                return 0.0;
            } elseif (parsedComplex["imaginary"] < 0.0) {
                return M_PI / -2;
            } else {
                return M_PI / 2;
            }
        } elseif (parsedComplex["real"] > 0.0) {
            return atan(parsedComplex["imaginary"] / parsedComplex["real"]);
        } elseif (parsedComplex["imaginary"] < 0.0) {
            return 0 - (M_PI - atan(abs(parsedComplex["imaginary"]) / abs(parsedComplex["real"])));
        } else {
            return M_PI - atan(parsedComplex["imaginary"] / abs(parsedComplex["real"]));
        }
    }


    /**
     * IMCONJUGATE
     *
     * Returns the complex conjugate of a complex number in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMCONJUGATE(complexNumber)
     *
     * @param    string        complexNumber    The complex number for which you want the conjugate.
     * @return    string
     */
    public static function imconjugate(var complexNumber)
    {
        var parsedComplex, tmp;
        
        let complexNumber = \ZExcel\Calculation\Functions::flattenSingleValue(complexNumber);

        let parsedComplex = self::parseComplex(complexNumber);

        if (parsedComplex["imaginary"] == 0.0) {
            return parsedComplex["real"];
        } else {
            let parsedComplex["imaginary"] = 0 - (double) parsedComplex["imaginary"];
            
            let tmp = self::CoMPLEX(
                parsedComplex["real"],
                parsedComplex["imaginary"],
                parsedComplex["suffix"]
            );
            
            return self::cleanComplex(tmp);
        }
    }


    /**
     * IMCOS
     *
     * Returns the cosine of a complex number in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMCOS(complexNumber)
     *
     * @param    string        complexNumber    The complex number for which you want the cosine.
     * @return    string|float
     */
    public static function imcos(complexNumber)
    {
        var parsedComplex;
        
        let complexNumber = \ZExcel\Calculation\Functions::flattenSingleValue(complexNumber);

        let parsedComplex = self::parseComplex(complexNumber);

        if (parsedComplex["imaginary"] == 0.0) {
            return cos(parsedComplex["real"]);
        } else {
            return self::iMCONJUGATE(
                self::CoMPLEX(
                    cos(parsedComplex["real"]) * cosh(parsedComplex["imaginary"]),
                    sin(parsedComplex["real"]) * sinh(parsedComplex["imaginary"]),
                    parsedComplex["suffix"]
                )
            );
        }
    }


    /**
     * IMSIN
     *
     * Returns the sine of a complex number in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMSIN(complexNumber)
     *
     * @param    string        complexNumber    The complex number for which you want the sine.
     * @return    string|float
     */
    public static function imsin(complexNumber)
    {
        var parsedComplex;
        
        let complexNumber = \ZExcel\Calculation\Functions::flattenSingleValue(complexNumber);

        let parsedComplex = self::parseComplex(complexNumber);

        if (parsedComplex["imaginary"] == 0.0) {
            return sin(parsedComplex["real"]);
        } else {
            return self::CoMPLEX(
                sin(parsedComplex["real"]) * cosh(parsedComplex["imaginary"]),
                cos(parsedComplex["real"]) * sinh(parsedComplex["imaginary"]),
                parsedComplex["suffix"]
            );
        }
    }


    /**
     * IMSQRT
     *
     * Returns the square root of a complex number in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMSQRT(complexNumber)
     *
     * @param    string        complexNumber    The complex number for which you want the square root.
     * @return    string
     */
    public static function imsqrt(complexNumber)
    {
        var parsedComplex, theta, d1, d2, r;
        
        let complexNumber = \ZExcel\Calculation\Functions::flattenSingleValue(complexNumber);

        let parsedComplex = self::parseComplex(complexNumber);

        let theta = self::iMARGUMENT(complexNumber);
        let d1 = cos(theta / 2);
        let d2 = sin(theta / 2);
        let r = sqrt(sqrt((parsedComplex["real"] * parsedComplex["real"]) + (parsedComplex["imaginary"] * parsedComplex["imaginary"])));

        if (parsedComplex["suffix"] == "") {
            return self::CoMPLEX(d1 * r, d2 * r);
        } else {
            return self::CoMPLEX(d1 * r, d2 * r, parsedComplex["suffix"]);
        }
    }


    /**
     * IMLN
     *
     * Returns the natural logarithm of a complex number in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMLN(complexNumber)
     *
     * @param    string        complexNumber    The complex number for which you want the natural logarithm.
     * @return    string
     */
    public static function imln(complexNumber)
    {
        var parsedComplex, logR, t;
        
        let complexNumber = \ZExcel\Calculation\Functions::flattenSingleValue(complexNumber);

        let parsedComplex = self::parseComplex(complexNumber);

        if ((parsedComplex["real"] == 0.0) && (parsedComplex["imaginary"] == 0.0)) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        let logR = log(sqrt((parsedComplex["real"] * parsedComplex["real"]) + (parsedComplex["imaginary"] * parsedComplex["imaginary"])));
        let t = self::imargument(complexNumber);

        if (parsedComplex["suffix"] == "") {
            return self::CoMPLEX(logR, t);
        } else {
            return self::CoMPLEX(logR, t, parsedComplex["suffix"]);
        }
    }


    /**
     * IMLOG10
     *
     * Returns the common logarithm (base 10) of a complex number in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMLOG10(complexNumber)
     *
     * @param    string        complexNumber    The complex number for which you want the common logarithm.
     * @return    string
     */
    public static function imlog10(complexNumber)
    {
        var parsedComplex;
        
        let complexNumber = \ZExcel\Calculation\Functions::flattenSingleValue(complexNumber);

        let parsedComplex = self::parseComplex(complexNumber);

        if ((parsedComplex["real"] == 0.0) && (parsedComplex["imaginary"] == 0.0)) {
            return \ZExcel\Calculation\Functions::NaN();
        } elseif ((parsedComplex["real"] > 0.0) && (parsedComplex["imaginary"] == 0.0)) {
            return log10(parsedComplex["real"]);
        }

        return call_user_func(["\\ZExcel\\Calculation\\Engineering", "IMPRODUCT"], log10(self::EULER), self::imln(complexNumber));
    }


    /**
     * IMLOG2
     *
     * Returns the base-2 logarithm of a complex number in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMLOG2(complexNumber)
     *
     * @param    string        complexNumber    The complex number for which you want the base-2 logarithm.
     * @return    string
     */
    public static function imlog2(complexNumber)
    {
        var parsedComplex;
        
        let complexNumber = \ZExcel\Calculation\Functions::flattenSingleValue(complexNumber);

        let parsedComplex = self::parseComplex(complexNumber);

        if ((parsedComplex["real"] == 0.0) && (parsedComplex["imaginary"] == 0.0)) {
            return \ZExcel\Calculation\Functions::NaN();
        } elseif ((parsedComplex["real"] > 0.0) && (parsedComplex["imaginary"] == 0.0)) {
            return log(parsedComplex["real"], 2);
        }

        return call_user_func(["\\ZExcel\\Calculation\\Engineering", "IMPRODUCT"], log(self::EULER, 2), self::imln(complexNumber));
    }


    /**
     * IMEXP
     *
     * Returns the exponential of a complex number in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMEXP(complexNumber)
     *
     * @param    string        complexNumber    The complex number for which you want the exponential.
     * @return    string
     */
    public static function imexp(complexNumber)
    {
        var parsedComplex, e, eX, eY;
        
        let complexNumber = \ZExcel\Calculation\Functions::flattenSingleValue(complexNumber);

        let parsedComplex = self::parseComplex(complexNumber);

        if ((parsedComplex["real"] == 0.0) && (parsedComplex["imaginary"] == 0.0)) {
            return "1";
        }

        let e = exp(parsedComplex["real"]);
        let eX = e * cos(parsedComplex["imaginary"]);
        let eY = e * sin(parsedComplex["imaginary"]);

        if (parsedComplex["suffix"] == "") {
            return self::CoMPLEX(eX, eY);
        } else {
            return self::CoMPLEX(eX, eY, parsedComplex["suffix"]);
        }
    }


    /**
     * IMPOWER
     *
     * Returns a complex number in x + yi or x + yj text format raised to a power.
     *
     * Excel Function:
     *        IMPOWER(complexNumber,realNumber)
     *
     * @param    string        complexNumber    The complex number you want to raise to a power.
     * @param    float        realNumber        The power to which you want to raise the complex number.
     * @return    string
     */
    public static function impower(var complexNumber, var realNumber)
    {
        var parsedComplex, r, rPower, theta;
        
        let complexNumber = \ZExcel\Calculation\Functions::flattenSingleValue(complexNumber);
        let realNumber    = \ZExcel\Calculation\Functions::flattenSingleValue(realNumber);

        if (!is_numeric(realNumber)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        let parsedComplex = self::parseComplex(complexNumber);

        let r = sqrt((parsedComplex["real"] * parsedComplex["real"]) + (parsedComplex["imaginary"] * parsedComplex["imaginary"]));
        let rPower = pow(r, realNumber);
        let theta = self::imargument(complexNumber) * realNumber;
        if (theta == 0) {
            return 1;
        } elseif (parsedComplex["imaginary"] == 0.0) {
            return self::CoMPLEX(rPower * cos(theta), rPower * sin(theta), parsedComplex["suffix"]);
        } else {
            return self::CoMPLEX(rPower * cos(theta), rPower * sin(theta), parsedComplex["suffix"]);
        }
    }


    /**
     * IMDIV
     *
     * Returns the quotient of two complex numbers in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMDIV(complexDividend,complexDivisor)
     *
     * @param    string        complexDividend    The complex numerator or dividend.
     * @param    string        complexDivisor        The complex denominator or divisor.
     * @return    string
     */
    public static function imdiv(complexDividend, complexDivisor)
    {
        var parsedComplexDividend, parsedComplexDivisor, d1, d2, d3, r, i;
        
        let complexDividend = \ZExcel\Calculation\Functions::flattenSingleValue(complexDividend);
        let complexDivisor  = \ZExcel\Calculation\Functions::flattenSingleValue(complexDivisor);

        let parsedComplexDividend = self::parseComplex(complexDividend);
        let parsedComplexDivisor = self::parseComplex(complexDivisor);

        if ((parsedComplexDividend["suffix"] != "") && (parsedComplexDivisor["suffix"] != "") && (parsedComplexDividend["suffix"] != parsedComplexDivisor["suffix"])) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        if ((parsedComplexDividend["suffix"] != "") && (parsedComplexDivisor["suffix"] == "")) {
            let parsedComplexDivisor["suffix"] = parsedComplexDividend["suffix"];
        }

        let d1 = (parsedComplexDividend["real"] * parsedComplexDivisor["real"]) + (parsedComplexDividend["imaginary"] * parsedComplexDivisor["imaginary"]);
        let d2 = (parsedComplexDividend["imaginary"] * parsedComplexDivisor["real"]) - (parsedComplexDividend["real"] * parsedComplexDivisor["imaginary"]);
        let d3 = (parsedComplexDivisor["real"] * parsedComplexDivisor["real"]) + (parsedComplexDivisor["imaginary"] * parsedComplexDivisor["imaginary"]);

        let r = (string) (d1 / d3);
        let i = (string) (d2 / d3);

        if ((double) i > 0.0) {
            return self::cleanComplex(r . "+" . i . parsedComplexDivisor["suffix"]);
        } elseif (i < 0.0) {
            return self::cleanComplex(r . i . parsedComplexDivisor["suffix"]);
        } else {
            return r;
        }
    }


    /**
     * IMSUB
     *
     * Returns the difference of two complex numbers in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMSUB(complexNumber1,complexNumber2)
     *
     * @param    string        complexNumber1        The complex number from which to subtract complexNumber2.
     * @param    string        complexNumber2        The complex number to subtract from complexNumber1.
     * @return    string
     */
    public static function imsub()
    {
        var args, complexNumber1, complexNumber2, parsedComplex1, parsedComplex2, d1, d2;
        
        let args = func_get_args();
        
        if (count(args) < 2) {
            throw new \Exception("Required 2 arguments");
        }
        
        let complexNumber1 = \ZEXcel\Calculation\Functions::flattenSingleValue(args[0]);
        let complexNumber2 = \ZEXcel\Calculation\Functions::flattenSingleValue(args[1]);

        let parsedComplex1 = self::parseComplex(complexNumber1);
        let parsedComplex2 = self::parseComplex(complexNumber2);

        if (((parsedComplex1["suffix"] != "") && (parsedComplex2["suffix"] != "")) && (parsedComplex1["suffix"] != parsedComplex2["suffix"])) {
            return \ZEXcel\Calculation\Functions::NaN();
        } elseif ((parsedComplex1["suffix"] == "") && (parsedComplex2["suffix"] != "")) {
            let parsedComplex1["suffix"] = parsedComplex2["suffix"];
        }

        let d1 = parsedComplex1["real"] - parsedComplex2["real"];
        let d2 = parsedComplex1["imaginary"] - parsedComplex2["imaginary"];

        return self::CoMPLEX(d1, d2, parsedComplex1["suffix"]);
    }


    /**
     * IMSUM
     *
     * Returns the sum of two or more complex numbers in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMSUM(complexNumber[,complexNumber[,...]])
     *
     * @param    string        complexNumber,...    Series of complex numbers to add
     * @return    string
     */
    public static function imsum()
    {
        var returnValue, activeSuffix, aArgs, arg, parsedComplex;
        
        let returnValue = self::parseComplex("0");
        let activeSuffix = "";

        // Loop through the arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        for arg in aArgs {
            let parsedComplex = self::parseComplex(arg);

            if (activeSuffix == "") {
                let activeSuffix = parsedComplex["suffix"];
            } elseif ((parsedComplex["suffix"] != "") && (activeSuffix != parsedComplex["suffix"])) {
                return \ZExcel\Calculation\Functions::VaLUE();
            }

            let returnValue["real"] = returnValue["real"] + parsedComplex["real"];
            let returnValue["imaginary"] = returnValue["imaginary"] + parsedComplex["imaginary"];
        }

        if (returnValue["imaginary"] == 0.0) {
            let activeSuffix = "";
        }
        return self::CoMPLEX(returnValue["real"], returnValue["imaginary"], activeSuffix);
    }


    /**
     * IMPRODUCT
     *
     * Returns the product of two or more complex numbers in x + yi or x + yj text format.
     *
     * Excel Function:
     *        IMPRODUCT(complexNumber[,complexNumber[,...]])
     *
     * @param    string        complexNumber,...    Series of complex numbers to multiply
     * @return    string
     */
    public static function improduct()
    {
        var returnValue, activeSuffix, aArgs, arg, parsedComplex, workValue;
        
        let returnValue = self::parseComplex("1");
        let activeSuffix = "";

        // Loop through the arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        
        for arg in aArgs {
            let parsedComplex = self::parseComplex(arg);

            let workValue = returnValue;
            
            if ((parsedComplex["suffix"] != "") && (activeSuffix == "")) {
                let activeSuffix = parsedComplex["suffix"];
            } elseif ((parsedComplex["suffix"] != "") && (activeSuffix != parsedComplex["suffix"])) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let returnValue["real"] = (workValue["real"] * parsedComplex["real"]) - (workValue["imaginary"] * parsedComplex["imaginary"]);
            let returnValue["imaginary"] = (workValue["real"] * parsedComplex["imaginary"]) + (workValue["imaginary"] * parsedComplex["real"]);
        }

        if (returnValue["imaginary"] == 0.0) {
            let activeSuffix = "";
        }
        return self::CoMPLEX(returnValue["real"], returnValue["imaginary"], activeSuffix);
    }


    /**
     *    DELTA
     *
     *    Tests whether two values are equal. Returns 1 if number1 = number2; returns 0 otherwise.
     *    Use this function to filter a set of values. For example, by summing several DELTA
     *    functions you calculate the count of equal pairs. This function is also known as the
     *    Kronecker Delta function.
     *
     *    Excel Function:
     *        DELTA(a[,b])
     *
     *    @param    float        a    The first number.
     *    @param    float        b    The second number. If omitted, b is assumed to be zero.
     *    @return    int
     */
    public static function delta(var a, var b = 0) -> int
    {
        let a = \ZExcel\Calculation\Functions::flattenSingleValue(a);
        let b = \ZExcel\Calculation\Functions::flattenSingleValue(b);

        return (int) (a == b);
    }


    /**
     *    GESTEP
     *
     *    Excel Function:
     *        GESTEP(number[,step])
     *
     *    Returns 1 if number >= step; returns 0 (zero) otherwise
     *    Use this function to filter a set of values. For example, by summing several GESTEP
     *    functions you calculate the count of values that exceed a threshold.
     *
     *    @param    float        number        The value to test against step.
     *    @param    float        step        The threshold value.
     *                                    If you omit a value for step, GESTEP uses zero.
     *    @return    int
     */
    public static function gestep(var number, var step = 0) -> int
    {
        let number = \ZExcel\Calculation\Functions::flattenSingleValue(number);
        let step   = \ZExcel\Calculation\Functions::flattenSingleValue(step);

        return (int) (number >= step);
    }

    public static function erfVal(double x) -> double
    {
        double sum, term, xsqr, j;
        
        if (abs(x) > 2.2) {
            return 1 - (double) self::erfcVal(x);
        }
        
        let sum = x;
        let term = x;
        let xsqr = (x * x);
        let j = 1;

        do {
            let term = term * (xsqr / j);
            let sum = sum - (term / (2 * j + 1));
            let j = j + 1;
            let term = term * (xsqr / j);
            let sum = sum + (term / (2 * j + 1));
            let j = j + 1;
            
            if (sum == 0.0) {
                break;
            }
        } while (abs(term * (1 / sum)) > \ZEXcel\Calculation\Functions::PRECISION);
        
        return (double) self::twoSqrtPi * sum;
    }


    /**
     *    ERF
     *
     *    Returns the error function integrated between the lower and upper bound arguments.
     *
     *    Note: In Excel 2007 or earlier, if you input a negative value for the upper or lower bound arguments,
     *            the function would return a #NUM! error. However, in Excel 2010, the function algorithm was
     *            improved, so that it can now calculate the function for both positive and negative ranges.
     *            PHPExcel follows Excel 2010 behaviour, and accepts nagative arguments.
     *
     *    Excel Function:
     *        ERF(lower[,upper])
     *
     *    @param    float        lower    lower bound for integrating ERF
     *    @param    float        upper    upper bound for integrating ERF.
     *                                If omitted, ERF integrates between zero and lower_limit
     *    @return    float
     */
    public static function erf(var lower, var upper = null)
    {
        let lower = \ZExcel\Calculation\Functions::flattenSingleValue(lower);
        let upper = \ZExcel\Calculation\Functions::flattenSingleValue(upper);

        if (is_numeric(lower)) {
            if (is_null(upper)) {
                return self::erfVal(lower);
            }
            
            if (is_numeric(upper)) {
                return self::erfVal(upper) - self::erfVal(lower);
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }

    private static function erfcVal(double x) -> double
    {
        double a, n, b, c, d, q1, q2, t;
        
        if (abs(x) < 2.2) {
            return 1 - (double) self::erfVal(x);
        }
        
        if (x < 0) {
            return 2 - (double) self::erfc(-x);
        }
        
        let a = 1;
        let n = 1;
        let b = x;
        let c = x;
        let d = (x * x) + 0.5;
        let q1 = b * (1 / d);
        let q2 = q1;
        let t = 0;
        
        do {
            let t = a * n + b * x;
            let a = b;
            let b = t;
            let t = c * n + d * x;
            let c = d;
            let d = t;
            let n = n + 0.5;
            let q1 = q2;
            let q2 = b * (1 / d);
            
            let t = ((double) abs(q1 - q2) / q2);
        } while (t > \ZEXcel\Calculation\Functions::PRECISION);
        
        return (double) self::oneSqrtPi * exp(-x * x) * q2;
    }


    /**
     *    ERFC
     *
     *    Returns the complementary ERF function integrated between x and infinity
     *
     *    Note: In Excel 2007 or earlier, if you input a negative value for the lower bound argument,
     *        the function would return a #NUM! error. However, in Excel 2010, the function algorithm was
     *        improved, so that it can now calculate the function for both positive and negative x values.
     *            PHPExcel follows Excel 2010 behaviour, and accepts nagative arguments.
     *
     *    Excel Function:
     *        ERFC(x)
     *
     *    @param    float    x    The lower bound for integrating ERFC
     *    @return    float
     */
    public static function erfc(x)
    {
        let x = \ZExcel\Calculation\Functions::flattenSingleValue(x);

        if (is_numeric(x)) {
            return self::erfcVal(x);
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     *    getConversionGroups
     *    Returns a list of the different conversion groups for UOM conversions
     *
     *    @return    array
     */
    public static function getConversionGroups()
    {
        var conversionUnit;
        array conversionGroups = [];
        
        for conversionUnit in self::conversionUnits {
            let conversionGroups[] = conversionUnit["Group"];
        }
        
        return array_merge(array_unique(conversionGroups), []);
    }


    /**
     *    getConversionGroupUnits
     *    Returns an array of units of measure, for a specified conversion group, or for all groups
     *
     *    @param    string    group    The group whose units of measure you want to retrieve
     *    @return    array
     */
    public static function getConversionGroupUnits(group = null)
    {
        var conversionUnit, conversionGroup;
        array conversionGroups = [];
        
        for conversionUnit, conversionGroup in self::conversionUnits {
            if ((is_null(group)) || (conversionGroup["Group"] == group)) {
                let conversionGroups[conversionGroup["Group"]][] = conversionUnit;
            }
        }
        
        return conversionGroups;
    }


    /**
     *    getConversionGroupUnitDetails
     *
     *    @param    string    group    The group whose units of measure you want to retrieve
     *    @return    array
     */
    public static function getConversionGroupUnitDetails(group = null)
    {
        var conversionUnit, conversionGroup;
        array conversionGroups = [];
        
        for conversionUnit, conversionGroup in self::conversionUnits {
            if ((is_null(group)) || (conversionGroup["Group"] == group)) {
                let conversionGroups[conversionGroup["Group"]][] = [
                    "unit": conversionUnit,
                    "description": conversionGroup["Unit Name"]
                ];
            }
        }
        return conversionGroups;
    }


    /**
     *    getConversionMultipliers
     *    Returns an array of the Multiplier prefixes that can be used with Units of Measure in CONVERTUOM()
     *
     *    @return    array of mixed
     */
    public static function getConversionMultipliers()
    {
        return self::conversionMultipliers;
    }


    /**
     *    CONVERTUOM
     *
     *    Converts a number from one measurement system to another.
     *    For example, CONVERT can translate a table of distances in miles to a table of distances
     *    in kilometers.
     *
     *    Excel Function:
     *        CONVERT(value,fromUOM,toUOM)
     *
     *    @param    float        value        The value in fromUOM to convert.
     *    @param    string        fromUOM    The units for value.
     *    @param    string        toUOM        The units for the result.
     *
     *    @return    float
     */
    public static function convertuom(value, fromUOM, toUOM)
    {
        var fromMultiplier, toMultiplier, unitGroup1, fromUOM, unitGroup2;
        
        let value   = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let fromUOM = \ZExcel\Calculation\Functions::flattenSingleValue(fromUOM);
        let toUOM   = \ZExcel\Calculation\Functions::flattenSingleValue(toUOM);

        if (!is_numeric(value)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let fromMultiplier = 1.0;
        
        if (isset(self::conversionUnits[fromUOM])) {
            let unitGroup1 = self::conversionUnits[fromUOM]["Group"];
        } else {
            let fromMultiplier = substr(fromUOM, 0, 1);
            let fromUOM = substr(fromUOM, 1);
            
            if (isset(self::conversionMultipliers[fromMultiplier])) {
                let fromMultiplier = self::conversionMultipliers[fromMultiplier]["multiplier"];
            } else {
                return \ZExcel\Calculation\Functions::Na();
            }
            
            if ((isset(self::conversionUnits[fromUOM])) && (self::conversionUnits[fromUOM]["AllowPrefix"])) {
                let unitGroup1 = self::conversionUnits[fromUOM]["Group"];
            } else {
                return \ZExcel\Calculation\Functions::Na();
            }
        }
        let value = value * fromMultiplier;

        let toMultiplier = 1.0;
        
        if (isset(self::conversionUnits[toUOM])) {
            let unitGroup2 = self::conversionUnits[toUOM]["Group"];
        } else {
            let toMultiplier = substr(toUOM, 0, 1);
            let toUOM = substr(toUOM, 1);
            
            if (isset(self::conversionMultipliers[toMultiplier])) {
                let toMultiplier = self::conversionMultipliers[toMultiplier]["multiplier"];
            } else {
                return \ZExcel\Calculation\Functions::Na();
            }
            
            if ((isset(self::conversionUnits[toUOM])) && (self::conversionUnits[toUOM]["AllowPrefix"])) {
                let unitGroup2 = self::conversionUnits[toUOM]["Group"];
            } else {
                return \ZExcel\Calculation\Functions::Na();
            }
        }
        
        if (unitGroup1 != unitGroup2) {
            return \ZExcel\Calculation\Functions::Na();
        }

        if ((fromUOM == toUOM) && (fromMultiplier == toMultiplier)) {
            // We"ve already factored fromMultiplier into the value, so we need to reverse it again
            return value / fromMultiplier;
        } elseif (unitGroup1 == "Temperature") {
            if ((fromUOM == "F") || (fromUOM == "fah")) {
                if ((toUOM == "F") || (toUOM == "fah")) {
                    return value;
                } else {
                    let value = ((value - 32) / 1.8);
                    if ((toUOM == "K") || (toUOM == "kel")) {
                        let value = value + 273.15;
                    }
                    return value;
                }
            } elseif (((fromUOM == "K") || (fromUOM == "kel")) && ((toUOM == "K") || (toUOM == "kel"))) {
                        return value;
            } elseif (((fromUOM == "C") || (fromUOM == "cel")) && ((toUOM == "C") || (toUOM == "cel"))) {
                    return value;
            }
            if ((toUOM == "F") || (toUOM == "fah")) {
                if ((fromUOM == "K") || (fromUOM == "kel")) {
                    let value = value - 273.15;
                }
                return (value * 1.8) + 32;
            }
            if ((toUOM == "C") || (toUOM == "cel")) {
                return value - 273.15;
            }
            return value + 273.15;
        }
        return (value * self::unitConversions[unitGroup1][fromUOM][toUOM]) / toMultiplier;
    }
}
