FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

COPY openapi2excel.sln ./
COPY src/openapi2excel/openapi2excel.core.csproj src/openapi2excel/
COPY src/openapi2excel.cli/openapi2excel.cli.csproj src/openapi2excel.cli/

RUN dotnet restore src/openapi2excel.cli/openapi2excel.cli.csproj

COPY src ./src

RUN dotnet publish src/openapi2excel.cli/openapi2excel.cli.csproj \
    -c Release \
    -f net8.0 \
    -o /app/publish \
    --no-restore

FROM mcr.microsoft.com/dotnet/runtime:8.0 AS runtime
WORKDIR /app

COPY --from=build /app/publish ./

ENTRYPOINT ["dotnet", "openapi2excel.dll"]
