## About 

Here's a silly example that's re-usable. You can fetch multiple json files using parameters. 

The tech tree table is:
```ts
let
    Source = #"Get Tables",
    Data   = Source{[ Name = "TechTree" ]},
    Final  = Table.FromRecords( Data[Json] )
in 
    Final
```

## With a base query

```ts
// Get Tables
let
    BuildUrl = ( params ) as text  =>
        let 
            template = "/ninmonkey/MicrotopiaData/refs/heads/main/export/#[version]/json/#[file]"
        in 
            Text.Format( template, params ),

            // techtree-techtree.json
    CurrentVersion = "v1.0.8",
    Items = {
        [
            Name = "TechTree",
            Config = [
                RelativePath = BuildUrl( [ version = CurrentVersion, file = "techtree-techtree.json" ] )  
            ]    
        ],
        [
            Name = "CrusherOutput",
            Config = [
                RelativePath = BuildUrl( [ version = CurrentVersion, file = "crusher-output.json" ] )  
            ]    
        ]
    },

    InvokeApi = ( options ) as record => [ 
        Content = Web.Contents( "https://raw.githubusercontent.com", options ),
        Json = Json.Document( Content ),
        Meta = Value.Metadata( Content ), 
        RawText = Text.FromBinary( Content ),
        AbsoluteUrl = Meta[Content.Uri](),
        StatusCode = Meta[Response.Status],
        Size = Binary.Length( Content ),
        InferContent = Binary.InferContentType( Content ), // it guessed a wrong type
        HasErrors = StatusCode <> 200
    ],
    Summary = [
        Items = Items,
        Responses = List.Transform( Items, (r) => [ Name = r[Name], Data = InvokeApi( r[Config] ) ] ),
        // Data = Table.FromRecords( Responses, null, MissingField.Error ),
        Data = Table.FromRecords( Responses )
    ],
    Data = Summary[Data],
    #"Expanded Data" = Table.ExpandRecordColumn(Data, "Data", {"Content", "Json", "Meta", "RawText", "AbsoluteUrl", "StatusCode", "Size", "InferContent", "HasErrors"}, {"Content", "Json", "Meta", "RawText", "AbsoluteUrl", "StatusCode", "Size", "InferContent", "HasErrors"})
in
    #"Expanded Data"
```
