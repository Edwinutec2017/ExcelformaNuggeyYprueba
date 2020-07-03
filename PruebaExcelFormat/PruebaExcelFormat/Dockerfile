FROM microsoft/dotnet:2.2-runtime AS base
WORKDIR /app

FROM microsoft/dotnet:2.2-sdk AS build
WORKDIR /src
COPY PruebaExcelFormat/PruebaExcelFormat.csproj PruebaExcelFormat/
RUN dotnet restore PruebaExcelFormat/PruebaExcelFormat.csproj
COPY . .
WORKDIR /src/PruebaExcelFormat
RUN dotnet build PruebaExcelFormat.csproj -c Release -o /app

FROM build AS publish
RUN dotnet publish PruebaExcelFormat.csproj -c Release -o /app

FROM base AS final
WORKDIR /app
COPY --from=publish /app .
ENTRYPOINT ["dotnet", "PruebaExcelFormat.dll"]
