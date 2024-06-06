# OpenAPI-2-Excel

<div align="center">
    <img src="assets/logo.png" width="250px">
</div>

<div align="center">

[![🚧 - Under Development](https://img.shields.io/badge/🚧-Under_Development-orange)](https://)
[![NuGet Package](https://img.shields.io/badge/.NET%20-8.0-blue.svg)](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)
[![NuGet Package](https://img.shields.io/badge/Nugets-2ea44f?logo=nuget)](https://www.nuget.org/packages/openapi2excel.cli/)
[![GitHub license](https://img.shields.io/badge/License-MIT-green.svg)](https://github.com/pszybiak/openapi-2-excel/blob/main/LICENSE.md)

</div>

Tool to generate Rest API specification in a MS Excel format - human friendly document from Swagger/OpenAPI spec in YAML or JSON. The result should be accessible to Business Analyst and software developers.

> \[!NOTE]
>
> This project is part of the ["100 Commits"](https://100commitow.pl/) competition, whose main purpose is is to develop an original Open Source project for 100 days.
>

## Installation

Download and install the one of the currently supported [.NET SDKs](https://www.microsoft.com/net/download). Once installed, run the following command:

```bash
dotnet tool install --global openapi2excel.cli
```

## Usage

<div align="center">
    <img src="assets/usage.png" width="90%">
</div>

Example
```text
  openapi2excel C:\openapi-spec.yml C:\openapi-spec.xlsx
```

## Result

To show how the application works, let's use the official example used on the [Swagger Editor website](https://editor.swagger.io/).

```
  openapi2excel https://raw.githubusercontent.com/swagger-api/swagger-petstore/master/src/main/resources/openapi.yaml C:\openapi.xlsx
```

The first tab is an information tab, presenting document details and a list of available operations.

<div align="center">
    <img src="assets/info_worksheet.png" width="90%">
</div>

The next tabs contain details of individual operation.

<div align="center">
    <img src="assets/operation_details.png" width="90%">
</div>



## Wrap Up

If you think the repository can be improved, please open a PR with any updates and submit any issues.

## Contribution

- Open a pull request with improvements
- Discuss ideas in issues

## License

[![GitHub license](https://img.shields.io/badge/License-MIT-green.svg)](https://github.com/pszybiak/openapi-2-excel/blob/main/LICENSE.md)