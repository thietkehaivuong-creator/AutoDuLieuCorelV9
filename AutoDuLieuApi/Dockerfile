FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

COPY ["AutoDuLieuApi/AutoDuLieuApi.csproj", "AutoDuLieuApi/"]
RUN dotnet restore "AutoDuLieuApi/AutoDuLieuApi.csproj"

COPY . .
WORKDIR "/src/AutoDuLieuApi"
RUN dotnet build "AutoDuLieuApi.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "AutoDuLieuApi.csproj" -c Release -o /app/publish /p:UseAppHost=false

FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS final
WORKDIR /app
COPY --from=publish /app/publish .

RUN mkdir -p /app/data/config /app/data/mappings

EXPOSE 5000
ENV ASPNETCORE_URLS=http://+:5000
ENV ASPNETCORE_ENVIRONMENT=Production

ENTRYPOINT ["dotnet", "AutoDuLieuApi.dll"]