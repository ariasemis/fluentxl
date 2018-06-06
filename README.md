# FluentXL

FluentXL is a .NET Standard library that allows you to work with Excel files in a fluent, declarative way.

## Usage

```c#

var document = Spreadsheet
    .New()
    .WithSheet(Specification
        .Sheet()
        .WithName("sheet 1")
        .WithRow(Specification
            .Row()
            .OnIndex(1)
            .WithCell(Specification
                .Cell()
                .OnColumn(1)
                .WithContent("Hello World!!"))));

using (var ms = new MemoryStream())
{
    document.WriteTo(ms);
}

```

_For more examples and usage, please refer to the [Wiki][wiki]._

## Installation

Currently there aren't binaries available to download, so in order to use the library you need to:

1. Clone the repository
2. Build it on release
3. Go to the bin folder and copy the FluentXL.dll library
4. Import it on your project

## Contributing

Contributions are more than welcome:

1. [Fork it][fork]
2. Create your feature branch `git checkout -b feature/fooBar`
3. Commit your changes `git commit -am 'Add some fooBar'`
4. Push to the branch `git push origin feature/fooBar`
5. Create a new Pull Request

## License

This project is licensed under the MPLv2 License - see the [LICENSE](LICENSE) file for details

<!-- links -->
[wiki]: https://github.com/ariasemis/fluentxl/wiki
[fork]: https://github.com/ariasemis/fluentxl/fork
