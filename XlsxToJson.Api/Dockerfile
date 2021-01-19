#See https://aka.ms/containerfastmode to understand how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/aspnet:5.0-buster-slim AS base
RUN apt-get update && apt-get install -y libgdiplus
WORKDIR /app
EXPOSE 80
EXPOSE 443

FROM mcr.microsoft.com/dotnet/sdk:5.0-buster-slim AS build
WORKDIR /src
COPY ["XlsxToJson.Api/XlsxToJson.Api.csproj", "XlsxToJson.Api/"]
RUN dotnet restore "XlsxToJson.Api/XlsxToJson.Api.csproj"
COPY . .
WORKDIR "/src/XlsxToJson.Api"
RUN dotnet build "XlsxToJson.Api.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "XlsxToJson.Api.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "XlsxToJson.Api.dll"]