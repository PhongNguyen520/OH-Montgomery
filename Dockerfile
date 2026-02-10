FROM mcr.microsoft.com/playwright/dotnet:v1.48.0-jammy AS base
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
COPY ["OH_Montgomery/OH_Montgomery.csproj", "OH_Montgomery/"]
RUN dotnet restore "OH_Montgomery/OH_Montgomery.csproj"
COPY . .
WORKDIR "/src/OH_Montgomery"
RUN dotnet build "OH_Montgomery.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "OH_Montgomery.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
# Install browsers (if strictly needed, though base image usually has them. Uncomment if issues arise)
# RUN pwsh bin/Debug/net8.0/playwright.ps1 install chromium
ENTRYPOINT ["dotnet", "OH_Montgomery.dll"]
