# VBA Variant Decimal specification

## Reference

https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary
https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/variant-data-type
https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/decimal-data-type
https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function

https://bytecomb.com/vba-internals-whats-in-a-variable/
https://bytecomb.com/vba-internals-decimal-variables-and-pointers-in-depth/
https://carolomeetsbarolo.wordpress.com/2014/06/09/the-decimal-data-sub-type-in-xl-vba/

## Variant (Numeric) format

    Variant store an integer and a byte array whose length is 14 bytes
    +--------+--------+--------+--------+--------+--------+--------+--------+
    |     vartype     |                       reserved                      |
    +--------+--------+--------+--------+--------+--------+--------+--------+
    +--------+--------+--------+--------+--------+--------+--------+--------+
    |                                  data                                 |
    +--------+--------+--------+--------+--------+--------+--------+--------+

    where
    * vartype is a 16-bit little-endian signed integer which represents VarType value.
    * data is a little-endian signed integer or a little-endian floating point number.

## Decimal format

    Decimal store an integer and a byte array whose length is 14 bytes
    +--------+--------+--------+--------+--------+--------+--------+--------+
    |     vartype     |  scale |  sign  |         data (high bytes)         |
    +--------+--------+--------+--------+--------+--------+--------+--------+
    +--------+--------+--------+--------+--------+--------+--------+--------+
    |                            data (low bytes)                           |
    +--------+--------+--------+--------+--------+--------+--------+--------+

    where
    * vartype is a 16-bit little-endian signed integer which represents VarType value 14 (vbDecimal).
    * scale is a 8-bit integer which represents a scaling factor
      (used to indicate either a whole number power of 10 to scale the integer down by,
       or that there should be no scaling)
    * sign is a 8-bit integer which represents a value indicating whether the decimal number is positive or negative.
    * data (high bytes) is a 32-bit little-endian unsigned integer.
    * data (low bytes) is a 64-bit little-endian unsigned integer.
