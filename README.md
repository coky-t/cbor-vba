# CBOR for VBA implementation

## Specification

[CBOR for VBA specification](src/cbor-vba-spec.md)

## Usage

### Prepare

Import [CBOR.bas](src/CBOR.bas)

### Serialization

```
    Dim Bytes() As Byte
    Bytes = CBOR.GetCborBytes(Data)
    ' do anything
```

### Deserialization

```
    Dim Value
    If CBOR.IsCborObject(CborBytes) Then
        Set Value = CBOR.GetValue(CborBytes)
        If TypeName(Value) = "Collection" Then
            ' do anything
        ElseIf TypeName(Value) = "Dictionary" Then
            ' do anything
        End If
    Else
        Value = CBOR.GetValue(CborBytes)
        ' do anything
    End If
```

## License

[MIT License](LICENSE)
