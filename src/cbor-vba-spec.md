# CBOR for VBA specification

## Reference

### CBOR specification

Concise Binary Object Representation (CBOR)
https://datatracker.ietf.org/doc/html/rfc8949
https://www.ietf.org/rfc/rfc8949.pdf

### VBA data types

https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary

## Serialization: type to format conversion for VBA

CBOR for VBA serializers convert VBA types into CBOR formats as following:

VBA types                        | major types    | output format
-------------------------------- | -------------- | ------------------------------------------------
0. Empty                         | 7: simple      | undefined
1. Null                          | 7: simple      | null
2. Integer                       | 0, 1: integer  | integer
3. Long                          | 0, 1: integer  | integer
4. Single                        | 7: float       | single-precision float
5. Double                        | 7: float       | double-precision float
6. Currency                      | -              | -
7. Date                          | -              | -
8. String                        | 3: text string | UTF-8 string
9. Object                        | -              | -
10. Error                        | -              | -
11. Boolean                      | 7: simple      | true, false
12. Variant                      | -              | -
13. DataObject                   | -              | -
14. Decimal                      | -              | -
17. Byte                         | 0: integer     | integer
20. LongLong                     | 0, 1: integer  | integer
36. User-defined Type            | -              | -
8192. Array (Single Dimension)   | 4: array       | array
8192. Array (Multiple Dimension) | -              | -
8209. Byte() (Single Dimension)  | 2: byte string | byte string

VBA types                  | major types  | output format
-------------------------- | ------------ | -------------
Nothing                    | 7: simple    | null
User-defined Class         | -            | -
Collection                 | 4: array     | array
Dictionary                 | 5: map       | map

### To Do - not yet implemented

VBA types                        | major types    | output format
-------------------------------- | -------------- | ------------------------------------------------
6. Currency                      | 6: tag         | decimal fraction
7. Date                          | 6: tag         | standard date/time string, epoch-based date/time
14. Decimal                      | 6: tag         | bignum, decimal fraction

## Deserialization: format to type conversion

CBOR for VBA deserializers convert CBOR formats into VBA types as following:

major type        | source formats               | VBA types
----------------- | ---------------------------- | --------------------------------------
0, 1: integer     | integer                      | Byte, Integer, Long, LongLong, Decimal
2: byte string    | byte string                  | Byte()
3: text string    | UTF-8 text string            | String
4: array          | array                        | Collection (or Array)
5: map            | map                          | Dictionary
6: tag            | 0. standard date/time string | -
6: tag            | 1. epoch-based date/time     | -
6: tag            | 2, 3. bignum                 | -
6: tag            | 4. decimal fraction          | -
6: tag            | 4. big float                 | -
6: tag            | (tag)                        | -
7: simple/float   | false and true               | Boolean
7: simple/float   | null                         | Null
7: simple/float   | undefined                    | Empty
7: simple/float   | half-precision float         | Single
7: simple/float   | single-precision float       | Single
7: simple/float   | double-precision float       | Double
7: simple/float   | break                        | -

### To Do - not yet implemented

major type        | source formats               | VBA types
----------------- | ---------------------------- | --------------------------------------
6: tag            | 0. standard date/time string | Date
6: tag            | 1. epoch-based date/time     | Date
6: tag            | 2, 3. bignum                 | Decimal
6: tag            | 4. decimal fraction          | Decimal

## Limitation - not supported

* 0x5b byte string (eight-byte uint64_t for n, and then n bytes follow)
* 0x7b UTF-8 string (eight-byte uint64_t for n, and then n bytes follow)
* 0x9b array (eight-byte uint64_t for n, and then n data items follow)
* 0xbb map (eight-byte uint64_t for n, and then n pairs of data items follow

## To Do - not yet implemented

### bignum

* 0xc2 unsigned bignum (data item "byte string" follows; see Section 3.4.3)
* 0xc3 negative bignum (data item "byte string" follows; see Section 3.4.3)
* 0xc4 decimal Fraction (data item "array" follows; see Section 3.4.4)
* 0xc5 bigfloat (data item "array" follows; see Section 3.4.4)

### break

* 0x5f byte string, byte strings follow, terminated by "break"
* 0x7f UTF-8 string, UTF-8 strings follow, terminated by "break"
* 0x9f array, data items follow, terminated by "break"
* 0xbf map, pairs of data items follow, terminated by "break"
* 0xff "break" stop code

### date/time

* 0xc0 text-based date/time (data item follows; see Section 3.4.1)
* 0xc1 epoch-based date/time (data item follows; see Section 3.4.2)

### map

* 4.2.1. Core Deteministic Encoding Requirements - The keys in every map MUST be sorted in the bytewise lexicographic order of their deterministic encodings.
* 4.2.3. Length-First Map Key Ordering

### tag

* 3.4.5.1. Encoded CBOR Data Item
* 3.4.5.2 Expecteed Later Encoding for CBOR-to-JSON Converters
* 3.4.5.3. Encoded Text
* 3.4.6. Self-Described CBOR

* 0xd5..0xd7 expected conversion (data item follows; see Section 3.4.5.2)
